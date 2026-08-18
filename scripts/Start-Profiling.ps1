<#
.SYNOPSIS
    Утренний прогон профайлера одной кнопкой: проверка среды, дамп исходников,
    сбор трассировки, отчёт - и откат всего, что было изменено.

.DESCRIPTION
    Сценарий связывает четыре скрипта задачи в один проход и на каждом шаге говорит,
    что делает и чем кончилось:

        1. Test-TraceEnvironment.ps1 - готова ли машина;
        2. служба инстанции          - запустить, если остановлена;
        3. Dump-AlSource.ps1         - обновить дамп исходников (.alsrc);
        4. Collect-AlTrace.ps1       - собрать трассировку;
        5. Rebuild-Report.ps1        - метрики, подсказки, HTML-отчёт;
        6. откат                     - вернуть службу в исходное состояние;
        7. открыть отчёт в браузере.

    Человек нужен ровно в одном месте: на шаге 4 сборщик просит выполнить сценарий в
    клиенте NAV и нажать Enter. Всё остальное идёт без вопросов.

    Порядок шагов не случаен. Дамп исходников обновляется ПОСЛЕ запуска службы и ДО
    сбора: номера строк в событиях трассировки верны только для того состояния объекта,
    в каком он скомпилирован в базе, и дамп должен быть снят с той же компиляции, что и
    трасса. Откат идёт до открытия отчёта, чтобы машина вернулась в исходное состояние
    даже если браузер не запустится.

    Любой шаг, кроме сбора, при ошибке останавливает прогон. Пустая трасса (код 8
    сборщика) прогон не валит: отчёт всё равно строится, но без таймингов - иначе
    непонятно, что именно не сошлось.

    Что меняется на машине и как возвращается:
      - служба инстанции: если была остановлена, запускается и в конце гасится обратно;
      - всё остальное (сессия ETW, EnableFullALFunctionTracing) - забота
        Collect-AlTrace.ps1, он печатает свой план и откатывает себя сам.

.PARAMETER Mode
    Full - полная трассировка операторов C/AL, нужны права администратора.
    Lite - только долгие SQL из журнала Application, прав не требует.

.PARAMETER ObjectType
    Тип объекта для отчёта: 1 Table, 3 Report, 5 Codeunit, 6 XMLport, 8 Page, 9 Query.
    По умолчанию 5 (Codeunit).

.PARAMETER ObjectId
    Номер объекта для отчёта. По умолчанию 80 (Sales-Post).

.PARAMETER ServerInstance
    Имя инстанции NAV. По умолчанию DynamicsNAV110.

.PARAMETER Server
    Экземпляр SQL Server для дампа исходников. По умолчанию localhost.

.PARAMETER Database
    База данных NAV. По умолчанию NAV либо значение LP_DATABASE.

.PARAMETER MaxSizeMB
    Размер кольцевого файла .etl, МБ. По умолчанию 1024.

.PARAMETER Calibrate
    Прогнать сценарий дважды и снять цену события трассировки.

.PARAMETER SkipDump
    Не обновлять дамп исходников: дамп уже снят с этой же компиляции.

.PARAMETER SkipEnvCheck
    Не проверять среду. Не рекомендуется: проверка стоит секунду и ловит
    ровно те грабли, из-за которых прогон оказывается пустым.

.PARAMETER NoOpen
    Не открывать отчёт в браузере.

.PARAMETER OutDir
    Каталог прогона. По умолчанию <корень репозитория>\out\run-<метка времени>.

.EXAMPLE
    .\Start-Profiling.ps1
    Обычный утренний прогон по кодюниту 80. Из консоли администратора.

.EXAMPLE
    .\Start-Profiling.ps1 -Mode Lite -ObjectType 5 -ObjectId 90
    Без прав администратора: долгие SQL плюс листинг кодюнита 90.

.EXAMPLE
    .\Start-Profiling.ps1 -Calibrate -MaxSizeMB 2048
    Полный прогон с замером цены события; сценарий выполняется дважды.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet('Full', 'Lite')]
    [string] $Mode           = 'Full',
    [int]    $ObjectType     = 5,
    [int]    $ObjectId       = 80,
    [string] $ServerInstance = 'DynamicsNAV110',
    [string] $Server         = 'localhost',
    [string] $Database       = $(if ($env:LP_DATABASE) { $env:LP_DATABASE } else { 'NAV' }),
    [ValidateRange(16, 16384)]
    [int]    $MaxSizeMB      = 1024,
    [switch] $Calibrate,
    [switch] $SkipDump,
    [switch] $SkipEnvCheck,
    [switch] $NoOpen,
    [string] $OutDir
)

$ErrorActionPreference = 'Stop'

$STEPS = 7
$svcName = 'MicrosoftDynamicsNavServer$' + $ServerInstance

function Write-StepHead {
    param([int] $No, [string] $Title)
    Write-Host ''
    Write-Host ('=== ШАГ {0}/{1}  {2} ' -f $No, $STEPS, $Title).PadRight(70, '=') -ForegroundColor Cyan
}

function Write-Ok   { param([string] $M) Write-Host ('  [ОК]   ' + $M) -ForegroundColor Green }
function Write-Warn { param([string] $M) Write-Host ('  [ВНИМ] ' + $M) -ForegroundColor Yellow }
function Write-Bad  { param([string] $M) Write-Host ('  [СБОЙ] ' + $M) -ForegroundColor Red }

# Дочерний скрипт запускается так, чтобы код возврата был достоверным: часть скриптов
# завершается без явного exit, и тогда $LASTEXITCODE хранил бы код от прошлого вызова.
function Invoke-Child {
    param(
        [Parameter(Mandatory)][string] $Path,
        [hashtable] $Arguments = @{}
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw ('Не найден скрипт ' + $Path) }
    $global:LASTEXITCODE = 0
    & $Path @Arguments
    $code = $LASTEXITCODE
    if ($null -eq $code) { $code = 0 }
    return [int]$code
}

# Линтер и построитель HTML здесь не перечислены: их зовёт Rebuild-Report.ps1,
# который отвечает за весь шаг «метрики - подсказки - отчёт» целиком.
$scripts = @{
    Env   = Join-Path $PSScriptRoot 'Test-TraceEnvironment.ps1'
    Dump  = Join-Path $PSScriptRoot 'Dump-AlSource.ps1'
    Trace = Join-Path $PSScriptRoot 'Collect-AlTrace.ps1'
    Build = Join-Path $PSScriptRoot 'Rebuild-Report.ps1'
}

if (-not $OutDir) {
    $taskRoot = Split-Path -Parent $PSScriptRoot          # scripts -> LineProfiler
    $OutDir   = Join-Path $taskRoot ('out\run-{0:yyyyMMdd-HHmmss}' -f (Get-Date))
}
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

# Свой протокол работы: без него сбой между шагами не оставляет следов —
# collect.log ведёт только сборщик, а консоль закрывается вместе с окном.
$transcriptPath = Join-Path $OutDir 'profiling.log'
try { Start-Transcript -Path $transcriptPath -Force | Out-Null; $transcriptOn = $true }
catch { $transcriptOn = $false }

$swAll        = [System.Diagnostics.Stopwatch]::StartNew()
$svcWasStopped = $false
$reportPath   = Join-Path $OutDir ('{0}_{1}.html' -f $ObjectType, $ObjectId)
$eventsPath   = Join-Path $OutDir 'events.tsv'
$exitCode     = 0
$results      = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param([string] $Step, [string] $Verdict, [string] $Note)
    [void]$results.Add([pscustomobject]@{ Step = $Step; Verdict = $Verdict; Note = $Note })
}

Write-Host ''
Write-Host ('ПРОГОН ПРОФАЙЛЕРА C/AL: режим {0}, инстанция {1}' -f $Mode, $ServerInstance) -ForegroundColor White
Write-Host ('Отчёт по объекту: тип {0}, номер {1}' -f $ObjectType, $ObjectId)
Write-Host ('Каталог прогона:  {0}' -f $OutDir)
Write-Host ''
Write-Host 'Что будет изменено на машине и как вернётся:'
Write-Host ('  служба {0}: если остановлена - запустим, в конце вернём как было' -f $svcName)
if ($Mode -eq 'Full') {
    Write-Host '  EnableFullALFunctionTracing и сессия ETW: их меняет и откатывает Collect-AlTrace.ps1'
    Write-Host '  (он печатает собственный план перед первым изменяющим действием)'
} else {
    Write-Host '  режим Lite: трассировка не включается, инстанция не перезапускается'
}
Write-Host ''
Write-Host 'Человек нужен один раз: на шаге 4 выполнить сценарий в NAV и нажать Enter.'

try {
    # ---- 1. проверка среды -------------------------------------------------
    Write-StepHead 1 'Проверка среды'
    if ($SkipEnvCheck) {
        Write-Warn 'пропущена по -SkipEnvCheck'
        Add-Result 'Проверка среды' 'ПРОПУСК' '-SkipEnvCheck'
    } else {
        $code = Invoke-Child -Path $scripts.Env -Arguments @{
            ServerInstance = $ServerInstance; Server = $Server; Database = $Database }
        if ($code -ne 0) {
            # в Lite половина проверок неприменима: там ни прав, ни рестартов не нужно
            if ($Mode -eq 'Lite') {
                Write-Warn 'есть блокирующие пункты, но режиму Lite они не мешают - идём дальше'
                Add-Result 'Проверка среды' 'ВНИМ' 'блокирующие пункты не мешают Lite'
            } else {
                Write-Bad 'среда не готова - см. строки с вердиктом НЕТ выше'
                Add-Result 'Проверка среды' 'СБОЙ' 'есть блокирующие пункты'
                throw 'среда не готова к сбору трассировки'
            }
        } else {
            Write-Ok 'машина готова'
            Add-Result 'Проверка среды' 'ОК' ''
        }
    }

    # ---- 2. служба инстанции ------------------------------------------------
    Write-StepHead 2 'Служба инстанции'
    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Warn ('служба {0} не найдена - сценарий выполнять негде' -f $svcName)
        Add-Result 'Служба' 'ВНИМ' 'службы нет'
    } elseif ($svc.Status -eq 'Running') {
        Write-Ok ('{0} уже запущена, состояние не меняем' -f $svcName)
        Add-Result 'Служба' 'ОК' 'была запущена'
    } else {
        Write-Host ('  {0}: {1} -> запускаем (в конце вернём в {1})' -f $svcName, $svc.Status)
        $svcWasStopped = $true
        Start-Service -Name $svcName
        $svc.WaitForStatus('Running', [TimeSpan]::FromMinutes(5))
        $svc.Refresh()
        Write-Ok ('{0} запущена' -f $svcName)
        Add-Result 'Служба' 'ОК' 'запущена, вернём в Stopped'
    }

    # ---- 3. дамп исходников -------------------------------------------------
    Write-StepHead 3 'Дамп исходников C/AL'
    if ($SkipDump) {
        Write-Warn 'пропущен по -SkipDump; номера строк должны быть от той же компиляции'
        Add-Result 'Дамп исходников' 'ПРОПУСК' '-SkipDump'
    } else {
        $code = Invoke-Child -Path $scripts.Dump -Arguments @{
            Server = $Server; Database = $Database; Quiet = $true }
        if ($code -ne 0) {
            Write-Bad 'дамп не снят'
            Add-Result 'Дамп исходников' 'СБОЙ' ('код ' + $code)
            throw 'не удалось обновить дамп исходников'
        }
        Write-Ok 'дамп обновлён, нумерация строк соответствует текущей компиляции'
        Add-Result 'Дамп исходников' 'ОК' ''
    }

    # ---- 4. сбор трассировки ------------------------------------------------
    Write-StepHead 4 'Сбор трассировки (нужен человек)'
    $traceArgs = @{
        Mode = $Mode; ServerInstance = $ServerInstance; OutDir = $OutDir; Yes = $true }
    if ($Mode -eq 'Full') {
        $traceArgs['MaxSizeMB'] = $MaxSizeMB
        if ($Calibrate) { $traceArgs['Calibrate'] = $true }
    } else {
        $traceArgs['Wait'] = $true
    }
    $code = Invoke-Child -Path $scripts.Trace -Arguments $traceArgs
    if ($code -eq 0) {
        Write-Ok 'трассировка собрана'
        Add-Result 'Сбор трассировки' 'ОК' ''
    } elseif ($code -eq 9) {
        # Откат у сборщика не прошёл: на сервере осталась включённая трассировка либо
        # неподнятая инстанция. Данные, возможно, целы, но разбираться надо с машиной,
        # и зелёным такой прогон объявлять нельзя.
        Write-Bad 'ОТКАТ НЕ УДАЛСЯ: машина осталась изменённой - см. restore.cmd в каталоге прогона'
        Add-Result 'Сбор трассировки' 'СБОЙ' 'код 9: откат не удался'
        $exitCode = 9
    } elseif ($code -eq 8) {
        # пустая или потерянная трасса - не повод бросать прогон: листинг всё равно нужен,
        # а причина уже напечатана сборщиком
        Write-Warn 'трасса пустая либо потеряна - отчёт будет без таймингов'
        Add-Result 'Сбор трассировки' 'ВНИМ' 'код 8: пусто либо потери'
        $exitCode = 8
    } else {
        Write-Bad ('сборщик вернул код ' + $code)
        Add-Result 'Сбор трассировки' 'СБОЙ' ('код ' + $code)
        throw ('сбор трассировки не удался, код ' + $code)
    }

    # ---- 5. метрики, подсказки, отчёт ---------------------------------------
    # Шаг целиком вынесен в Rebuild-Report.ps1. Он же умеет пересобирать отчёт по
    # готовой трассировке — благодаря этому шаг проверяется без сбора и без
    # человека у клиента NAV, а раньше ошибку в нём было видно только по итогам
    # полного прогона.
    Write-StepHead 5 'Метрики и отчёт'
    $linesPath = Join-Path $OutDir 'lines.tsv'
    $hintsPath = Join-Path $OutDir 'hints.tsv'
    $code = Invoke-Child -Path $scripts.Build -Arguments @{
        RunDir = $OutDir; ObjectType = $ObjectType; ObjectId = $ObjectId; OutFile = $reportPath }

    if (Test-Path -LiteralPath $linesPath) {
        $nLines = [math]::Max(0, ([System.IO.File]::ReadAllLines($linesPath).Length - 1))
        Add-Result 'Метрики' 'ОК' ('строк ' + $nLines)
    } else {
        Add-Result 'Метрики' 'ВНИМ' 'нет lines.tsv'
    }
    if (Test-Path -LiteralPath $hintsPath) {
        Add-Result 'Подсказки' 'ОК' (Split-Path -Leaf $hintsPath)
    } else {
        Add-Result 'Подсказки' 'ВНИМ' 'не построены'
    }

    if ($code -eq 9) {
        # Гейт валидности забраковал прогон: HTML есть, но числа в нём недостоверны.
        Write-Bad 'прогон признан невалидным - числам в отчёте верить нельзя, разбор в сводке выше'
        Add-Result 'Отчёт' 'СБОЙ' 'код 9: прогон невалиден'
        $exitCode = 9
    }
    elseif ($code -eq 8) {
        # В трассировке есть операторы, но не этого объекта: в клиенте выполнен другой
        # сценарий. Трасса вовсе без операторов (облегчённый сбор) сюда не попадает -
        # там код 0, и это норма по построению режима.
        Write-Bad ('в трассировке нет операторов объекта {0}/{1}' -f $ObjectType, $ObjectId)
        Add-Result 'Отчёт' 'ВНИМ' 'без таймингов'
        if ($exitCode -eq 0) { $exitCode = 8 }
    }
    elseif ($code -ne 0 -or -not (Test-Path -LiteralPath $reportPath)) {
        Write-Bad 'отчёт не построен'
        Add-Result 'Отчёт' 'СБОЙ' ('код ' + $code)
        throw 'не удалось построить отчёт'
    }
    else {
        Write-Ok ('отчёт: ' + $reportPath)
        Add-Result 'Отчёт' 'ОК' (Split-Path -Leaf $reportPath)
    }
}
catch {
    Write-Host ''
    Write-Bad $_.Exception.Message
    if ($exitCode -eq 0) { $exitCode = 1 }
}
finally {
    if ($transcriptOn) { try { Stop-Transcript | Out-Null } catch {} }
    # ---- 6. откат -----------------------------------------------------------
    # Сессию ETW, флаг трассировки и перезапуск инстанции откатывает сам сборщик -
    # у него это в собственном finally. Здесь возвращается ровно то, что менял этот
    # скрипт: служба, поднятая на шаге 2.
    Write-StepHead 6 'Откат'
    if ($svcWasStopped) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Stopped') {
                Stop-Service -Name $svcName -Force
                $svc.WaitForStatus('Stopped', [TimeSpan]::FromMinutes(5))
            }
            Write-Ok ('{0} остановлена - как было до прогона' -f $svcName)
            Add-Result 'Откат службы' 'ОК' 'вернули в Stopped'
        }
        catch {
            Write-Bad ('не удалось остановить {0}: {1}' -f $svcName, $_.Exception.Message)
            Add-Result 'Откат службы' 'СБОЙ' $_.Exception.Message
            if ($exitCode -eq 0) { $exitCode = 1 }
        }
    } else {
        Write-Ok 'этот скрипт состояние службы не менял - возвращать нечего'
    }
    if ($Mode -eq 'Full') {
        Write-Host '  трассировку и сессию ETW откатил Collect-AlTrace.ps1 (его блок finally)'
    }
}

# ---- 7. открыть отчёт ------------------------------------------------------
Write-StepHead 7 'Отчёт'
if ((Test-Path -LiteralPath $reportPath) -and -not $NoOpen) {
    try {
        Start-Process -FilePath $reportPath | Out-Null
        Write-Ok 'открыт в браузере'
    }
    catch { Write-Warn ('открыть не удалось: ' + $_.Exception.Message) }
} elseif (Test-Path -LiteralPath $reportPath) {
    Write-Ok ('готов: ' + $reportPath)
} else {
    Write-Warn 'отчёта нет'
}

# ---- итог ------------------------------------------------------------------
$swAll.Stop()
Write-Host ''
Write-Host ('{0,-22} {1,-8} {2}' -f 'ШАГ', 'ИТОГ', 'ПРИМЕЧАНИЕ')
Write-Host (('-' * 22) + ' ' + ('-' * 8) + ' ' + ('-' * 38))
foreach ($r in $results) {
    $color = 'Green'
    if     ($r.Verdict -eq 'СБОЙ') { $color = 'Red' }
    elseif ($r.Verdict -ne 'ОК')   { $color = 'Yellow' }
    Write-Host ('{0,-22} ' -f $r.Step) -NoNewline
    Write-Host ('{0,-8} ' -f $r.Verdict) -NoNewline -ForegroundColor $color
    Write-Host $r.Note
}
Write-Host (('-' * 22) + ' ' + ('-' * 8) + ' ' + ('-' * 38))
Write-Host ('Каталог прогона: {0}' -f $OutDir)
Write-Host ('Время прогона:   {0:N1} мин' -f $swAll.Elapsed.TotalMinutes)
if ($exitCode -eq 0) {
    Write-Host 'Прогон завершён.' -ForegroundColor Green
} elseif ($exitCode -eq 8) {
    Write-Host 'Прогон завершён с оговоркой: трасса пустая либо потеряна (см. шаг 4).' -ForegroundColor Yellow
} elseif ($exitCode -eq 9) {
    Write-Host 'Прогон завершён КРАСНЫМ. Смотрите строки СБОЙ выше: либо откат на сервере' -ForegroundColor Red
    Write-Host 'не удался и машина осталась изменённой, либо гейт забраковал сам прогон.' -ForegroundColor Red
} else {
    Write-Host 'Прогон прерван. Разбор - в collect.log каталога прогона.' -ForegroundColor Red
}

exit $exitCode
