<#
.SYNOPSIS
    Пересборка отчёта по уже собранной трассировке: метрики, подсказки, HTML.

.DESCRIPTION
    Это шаг «метрики и отчёт» одной кнопки (Start-Profiling.ps1), вынесенный
    отдельно. Причина простая: сбор трассировки требует человека у клиента NAV,
    поэтому шаг, который идёт ПОСЛЕ сбора, иначе нельзя проверить, не гоняя весь
    прогон целиком. Дважды оказывалось, что ломался именно он.

    Второе назначение — переcборка. Правило линтера подкрутили, вёрстку отчёта
    поправили: трасса та же, гонять сценарий заново незачем.

    Порядок: события -> калибровка нумерации -> метрики строк (lines.tsv) ->
    подсказки (hints.tsv) -> HTML. Подсказки необязательны: если линтер упал,
    отчёт всё равно строится, только без колонки подсказок.

    ВАЖНО: дамп исходников (.alsrc) должен быть снят с той же компиляции, что и
    трасса, иначе номера строк разъедутся. Если объект с тех пор перекомпилировали,
    пересобирать отчёт по старой трассе нельзя — нужен новый сбор. Косвенный
    признак расхождения виден в проценте совпадений при калибровке.

.PARAMETER RunDir
    Каталог прогона: в нём лежит events.tsv, туда же кладутся lines.tsv,
    hints.tsv и HTML.

.PARAMETER ObjectType
    Тип объекта: 1 Table, 3 Report, 5 Codeunit, 6 XMLport, 8 Page, 9 Query.

.PARAMETER ObjectId
    Номер объекта.

.PARAMETER EventsFile
    Файл событий. По умолчанию <RunDir>\events.tsv.

.PARAMETER OutFile
    HTML-отчёт. По умолчанию <RunDir>\<тип>_<номер>.html.

.PARAMETER SourceRoot
    Каталог дампа исходников. По умолчанию <корень репозитория>\.alsrc.

.PARAMETER NoHints
    Не звать линтер: только метрики и листинг с таймингами.

.OUTPUTS
    Код возврата: 0 — отчёт построен; 8 — в трассировке есть операторы, но не этого
    объекта, то есть в клиенте выполнен другой сценарий (отчёт построен, без таймингов);
    9 — отчёт построен, но гейт валидности прогона его забраковал: числам верить нельзя;
    1 — построить отчёт не удалось.

    Трасса вовсе без событий уровня оператора (облегчённый сбор) — не ошибка: отчёт
    строится без таймингов, код 0.

.EXAMPLE
    .\Rebuild-Report.ps1 -RunDir ..\out\run-20260814-102121 -ObjectType 3 -ObjectId 12436
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $RunDir,
    [Parameter(Mandatory)][int]    $ObjectType,
    [Parameter(Mandatory)][int]    $ObjectId,
    [string] $EventsFile,
    [string] $OutFile,
    [string] $SourceRoot,
    [switch] $NoHints,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Lib-AlListing.ps1')
. (Join-Path $PSScriptRoot 'Lib-AlTrace.ps1')

function Write-Note { param([string]$Text) if (-not $Quiet) { Write-Host ('   ' + $Text) } }
function Write-Warn2 { param([string]$Text) Write-Host ('   ' + $Text) -ForegroundColor Yellow }

if (-not (Test-Path -LiteralPath $RunDir)) { throw "Нет каталога прогона: $RunDir" }
$RunDir = (Resolve-Path -LiteralPath $RunDir).Path
if (-not $EventsFile) { $EventsFile = Join-Path $RunDir 'events.tsv' }
if (-not $OutFile)    { $OutFile    = Join-Path $RunDir ('{0}_{1}.html' -f $ObjectType, $ObjectId) }

$linesPath = Join-Path $RunDir 'lines.tsv'
$hintsPath = Join-Path $RunDir 'hints.tsv'
$lintPath  = Join-Path $PSScriptRoot 'Invoke-PerfLint.ps1'
$repPath   = Join-Path $PSScriptRoot 'Build-Report.ps1'

$reportArgs = @{ ObjectType = $ObjectType; ObjectId = $ObjectId; OutFile = $OutFile }
if ($SourceRoot) { $reportArgs['SourceRoot'] = $SourceRoot }
$exitCode = 0

if (Test-Path -LiteralPath $EventsFile) {
    $listArgs = @{ ObjectType = $ObjectType; ObjectId = $ObjectId }
    if ($SourceRoot) { $listArgs['SourceRoot'] = $SourceRoot }
    $listing = @(Get-AlListing @listArgs)
    $events  = Import-AlTraceEvents -Path $EventsFile

    # Проверяем, что в трассировке вообще есть операторы нужного объекта.
    # Параметр -ObjectId задаёт лишь то, ПО ЧЕМУ строить отчёт; выполнить
    # соответствующий сценарий в клиенте должен человек. Если он выполнил
    # другое, отчёт вышел бы пустым и молча — вместо этого говорим прямо.
    $anyStmt = @($events | Where-Object { $_.Kind -eq 'Stmt' })
    $own = @($anyStmt | Where-Object { $_.ObjectType -eq $ObjectType -and $_.ObjectId -eq $ObjectId })

    if ($anyStmt.Count -eq 0) {
        # Облегчённый сбор шлёт только события SQL, без номеров строк: операторов в такой
        # трассе нет ПО ПОСТРОЕНИЮ. Раньше это валилось в общую ветку «нет операторов
        # объекта» - красная строка и код 8 при идеально отработавшем сборе, то есть код 8
        # переставал отличать норму от настоящего сбоя.
        Write-Warn2 'в трассировке нет событий уровня оператора - похоже на облегчённый сбор'
        Write-Warn2 'отчёт будет без таймингов; для построчного времени нужен полный режим'
    }
    elseif ($own.Count -eq 0) {
        Write-Host ('   в трассировке нет ни одного оператора объекта {0}/{1}' -f $ObjectType, $ObjectId) -ForegroundColor Red
        Write-Warn2 'Похоже, в клиенте выполнен другой сценарий. В трассировке есть:'
        $events | Where-Object { $_.Kind -eq 'Stmt' } |
            Group-Object ObjectType, ObjectId |
            Sort-Object Count -Descending | Select-Object -First 8 |
            ForEach-Object {
                Write-Host ('     объект {0,-12} операторов {1}' -f $_.Name, $_.Count) -ForegroundColor Yellow
            }
        $exitCode = 8
    }
    else {
        $off = Resolve-AlLineOffset -Events $events -Listing $listing `
                                    -ObjectType $ObjectType -ObjectId $ObjectId
        Write-Note ('калибровка нумерации: смещение {0}, совпадений {1} %' -f $off.Offset, $off.MatchPct)

        $measureArgs = @{
            Events = $events; LineOffset = $off.Offset; Listing = $listing
            ObjectType = $ObjectType; ObjectId = $ObjectId
        }
        if (-not $off.Ok) {
            Write-Warn2 'ВНИМАНИЕ: калибровка не сошлась, номера строк могут быть смещены'
            if ($off.Message) { Write-Warn2 $off.Message }
            # Резервный маппинг строится ровно для этого случая, и его собственное
            # сообщение велит его передать - а не передавал его никто. Берём ТОЛЬКО его:
            # строка без текстового совпадения пусть лучше выпадет из отчёта, чем ляжет
            # на соседнюю. Выпадение видно по «учтено % времени» в сводке ниже,
            # а подмену строки не видно ничем.
            if ($off.Fallback -and $off.Fallback.Count -gt 0) {
                $measureArgs['FallbackMap'] = $off.Fallback
                $measureArgs['LineSource']  = 'Fallback'
                Write-Warn2 ('перехожу на сопоставление по тексту: однозначных строк {0}' -f $off.Fallback.Count)
            }
            else {
                Write-Warn2 'резервного маппинга нет - считаю по смещению как есть'
            }
        }
        $metrics = Measure-AlLines @measureArgs
        Export-AlLineMetrics -Lines $metrics -Path $linesPath | Out-Null
        Write-Note ('метрики: строк {0}' -f @($metrics).Count)

        # Гейт валидности прогона. Шапка Lib-AlTrace.ps1 объявляет «учтено времени»
        # инвариантом, который гейтит прогон, - но в конвейере сводку не звал никто, и
        # развалившееся дерево вызовов давало внешне нормальный HTML. События подаются
        # ЦЕЛИКОМ: нумерация записей канала сквозная, на отфильтрованном наборе разрывы
        # были бы ложными.
        $summary = Get-AlRunSummary -Events $events
        if (-not $Quiet) { Write-AlRunSummary -Summary $summary }
        if (-not $summary.IsValid) {
            if ($Quiet) { foreach ($b in $summary.Blockers) { Write-Host ('   ' + $b) -ForegroundColor Red } }
            Write-Host '   прогон признан невалидным: числам в отчёте верить нельзя' -ForegroundColor Red
            $exitCode = 9
        }
        $reportArgs['MetricsFile'] = $linesPath

        # Подсказки — вещь полезная, но не обязательная: если линтер упал, отчёт
        # с метриками всё равно должен быть построен, поэтому его сбой ловится тут.
        if (-not $NoHints) {
            $lintOk = $false
            try {
                $global:LASTEXITCODE = 0
                & $lintPath -ObjectType $ObjectType -ObjectId $ObjectId `
                            -MetricsFile $linesPath -OutFile $hintsPath -Quiet
                $lintOk = ($LASTEXITCODE -eq 0 -or $null -eq $LASTEXITCODE)
            }
            catch {
                Write-Host ('   линтер упал: ' + $_.Exception.Message) -ForegroundColor DarkYellow
            }
            if ($lintOk -and (Test-Path -LiteralPath $hintsPath)) {
                $reportArgs['HintsFile'] = $hintsPath
                Write-Note 'подсказки посчитаны'
            } else {
                Write-Host '   подсказки пропущены' -ForegroundColor DarkYellow
            }
        }
    }
}
else {
    Write-Warn2 ('нет файла событий {0} — отчёт будет без таймингов' -f $EventsFile)
}

$global:LASTEXITCODE = 0
& $repPath @reportArgs
$code = $LASTEXITCODE
if ($null -eq $code) { $code = 0 }
if ($code -ne 0 -or -not (Test-Path -LiteralPath $OutFile)) {
    Write-Host '   отчёт не построен' -ForegroundColor Red
    exit 1
}
exit $exitCode