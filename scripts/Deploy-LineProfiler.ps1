#requires -Version 7
<#
.SYNOPSIS
    Выкладка профайлера на стенд: сборка приёмника, объекты C/AL, проверка результата.

.DESCRIPTION
    Собирает C#-приёмник, кладёт его в Add-ins службы NST, импортирует и компилирует
    объекты задачи, после чего ПРОВЕРЯЕТ, что каждый объект действительно помечен
    скомпилированным в dbo.[Object].

    Порядок выкладки сборки продиктован тем, что запущенная служба держит файл:
    остановить NST -> копировать -> запустить NST -> почистить кэш C/SIDE -> компилировать.
    Кэш C/SIDE (%TEMP%\Microsoft Dynamics NAV\Add-Ins\<платформа>\<папка>_<версия>) обязателен
    к удалению: он привязан к имени папки и версии сборки, а версию мы держим неизменной,
    иначе объявление DotNet в C/AL пришлось бы править на каждую пересборку.

    Откат: перед копированием сохраняется прежняя сборка; при ошибке копирования или
    при неудачном старте службы она возвращается на место.

    ВАЖНО (грабли, ради которых и написан скрипт): импорт объекта СБРАСЫВАЕТ его флаг
    Compiled в 0. Если после импорта скомпилировать не все объекты, а часть, то на
    первом же обращении к недокомпилированной таблице сервер отвечает
    "Объект метаданных Table N не найден", причём в лог finsql ничего не пишет.
    Поэтому компилируются ВСЕ объекты из файла задачи, и результат сверяется с базой.

.PARAMETER SkipDll
    Не пересобирать и не выкладывать приёмник (служба не перезапускается).
    Быстрый путь, когда правился только C/AL.

.PARAMETER SkipCal
    Только сборка и выкладка приёмника, объекты не трогать.

.EXAMPLE
    pwsh scripts/Deploy-LineProfiler.ps1
    Полная выкладка: приёмник + объекты.

.EXAMPLE
    pwsh scripts/Deploy-LineProfiler.ps1 -SkipDll
    Только объекты C/AL, служба не перезапускается.
#>
[CmdletBinding()]
param(
    [switch]$SkipDll,
    [switch]$SkipCal,
    [string]$Database = $(if ($env:LP_DATABASE) { $env:LP_DATABASE } else { 'NAV' }),
    [string]$Server   = 'localhost'
)

$ErrorActionPreference = 'Stop'

$taskDir   = Split-Path $PSScriptRoot -Parent
$srcCs     = Join-Path $taskDir 'src\AlLineProfiler.cs'
$binDll    = Join-Path $taskDir 'bin\AlLineProfiler.dll'
$taskTxt   = Join-Path $taskDir 'LineProfiler.txt'
$taskCp866 = Join-Path $taskDir 'LineProfiler.cp866.txt'

$navRoot   = 'C:\Program Files\Microsoft Dynamics NAV\110\Service'
$addInDir  = Join-Path $navRoot 'Add-ins\LineProfiler'
$traceRef  = Join-Path $navRoot 'Add-ins\NavEtwReceiver\Microsoft.Diagnostics.Tracing.TraceEvent.dll'
$finsql    = 'C:\Program Files (x86)\Microsoft Dynamics NAV\110\RoleTailored Client\finsql.exe'
$csc       = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$service   = 'MicrosoftDynamicsNavServer$DynamicsNAV110'
$platform  = '11.0.36462.0'
$logDir    = Join-Path $env:TEMP 'LineProfiler-deploy'

$typeRank = @{ 'Table' = 1; 'Codeunit' = 2; 'Report' = 3; 'XMLport' = 4; 'Query' = 5; 'Page' = 6; 'MenuSuite' = 7 }
$typeCode = @{ 'Table' = 1; 'Report' = 3; 'Codeunit' = 5; 'XMLport' = 6; 'MenuSuite' = 7; 'Page' = 8; 'Query' = 9 }

$errors = @()
function Fail([string]$m) { $script:errors += $m; Write-Host "ОШИБКА: $m" -ForegroundColor Red }
function Step([string]$m) { Write-Host "  $m" }

# Флаг Compiled читается из базы, а не из вывода finsql: finsql про пустой фильтр
# и про несостоявшуюся компиляцию молчит и возвращает нулевой код возврата.
function Get-CompiledFlag($o, [string]$srv, [string]$db) {
    $conn = [System.Data.SqlClient.SqlConnection]::new(
        "Server=$srv;Database=$db;Integrated Security=True;TrustServerCertificate=True")
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT [Compiled] FROM dbo.[Object] WHERE [Type]=$($typeCode[$o.Type]) AND [ID]=$($o.Id)"
        $v = $cmd.ExecuteScalar()
        if ($null -eq $v -or $v -is [DBNull]) { return $null }
        return [int]$v
    } finally { $conn.Close() }
}

function Test-Compiled($o, [string]$srv, [string]$db) {
    return ((Get-CompiledFlag $o $srv $db) -eq 1)
}

# Служба переходит в Running задолго до того, как начинает обслуживать базу: сначала она
# компилирует бизнес-сборки (на полном наборе объектов это минуты). Компиляция C/AL,
# запущенная в это окно, падает с "Для этой базы данных недоступен экземпляр NAV Server",
# а объект с ссылкой DotNet вдобавок не может разрешить сборку из Add-ins.
# Отметка о готовности пишется в Operational, а НЕ в Admin (в Admin её нет вовсе).
function Wait-NstReady([datetime]$since, [int]$timeoutSec = 420) {
    $logs = @('Microsoft-DynamicsNAV-Server/Operational', 'Microsoft-DynamicsNAV-Server/Admin')
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ((Get-Date) -lt $deadline) {
        foreach ($log in $logs) {
            try {
                $ready = Get-WinEvent -FilterHashtable @{ LogName = $log; StartTime = $since } -ErrorAction Stop |
                    Where-Object { $_.Message -match 'completed configuration and is ready' }
                if ($ready) { return $true }
            } catch {}   # журнала может не быть или в нём пока пусто - это не отказ
        }
        Start-Sleep -Seconds 5
    }
    return $false
}

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

# ---------------------------------------------------------------- приёмник ETW

if (-not $SkipDll) {
    Write-Host 'Приёмник ETW' -ForegroundColor Cyan

    # На другой сборке платформы приёмник ETW от NAV может лежать в другой подпапке
    # (а на целевой базе поставочного может не быть вовсе - зато есть чужой, принесённый
    # сторонним профайлером). Берём любую найденную, старшую по версии.
    if (-not (Test-Path $traceRef)) {
        $found = Get-ChildItem (Join-Path $navRoot 'Add-ins') -Recurse -Filter 'Microsoft.Diagnostics.Tracing.TraceEvent.dll' -ErrorAction SilentlyContinue |
            Sort-Object { [version][Diagnostics.FileVersionInfo]::GetVersionInfo($_.FullName).FileVersion } -Descending |
            Select-Object -First 1
        if ($found) {
            $traceRef = $found.FullName
            Step "TraceEvent взят из $($found.Directory.Name)"
        }
    }

    foreach ($p in @($csc, $traceRef, $srcCs)) {
        if (-not (Test-Path $p)) { throw "не найдено: $p" }
    }

    # bin/ лежит в .gitignore, поэтому в свежем клоне его нет, а csc при отсутствии
    # каталога падает на генерации ресурса Win32 - по сообщению причина не читается.
    $binDir = Split-Path $binDll -Parent
    if (-not (Test-Path $binDir)) { New-Item -ItemType Directory -Path $binDir | Out-Null }

    & $csc /nologo /target:library /optimize+ /out:$binDll /reference:"$traceRef" $srcCs
    if ($LASTEXITCODE -ne 0) { throw "csc вернул $LASTEXITCODE" }
    Step "собрано: $([int]((Get-Item $binDll).Length)) байт"

    if (-not (Test-Path $addInDir)) { New-Item -ItemType Directory -Path $addInDir | Out-Null }
    $target = Join-Path $addInDir 'AlLineProfiler.dll'
    $backup = Join-Path $logDir 'AlLineProfiler.rollback.dll'
    $hadOld = Test-Path $target
    if ($hadOld) { Copy-Item $target $backup -Force }

    $svcWasRunning = (Get-Service $service).Status -eq 'Running'
    if ($svcWasRunning) {
        Step 'останов NST (файл держит служба)'
        Stop-Service $service -Force -Confirm:$false
    }

    # Служба отпускает файл не в момент перехода в Stopped, а когда процесс действительно
    # завершился, поэтому ждём именно освобождения файла, а не статуса службы.
    $free = -not $hadOld
    for ($i = 0; -not $free -and $i -lt 40; $i++) {
        try { $fs = [IO.File]::Open($target, 'Open', 'ReadWrite', 'None'); $fs.Close(); $free = $true }
        catch { Start-Sleep -Milliseconds 1500 }
    }

    $copied = $false
    if ($free) {
        try { Copy-Item $binDll $target -Force; $copied = $true; Step "выложено: $target" }
        catch { Fail "копирование не удалось: $($_.Exception.Message)" }
    } else {
        Fail 'сборка в Add-ins осталась занятой, копирование пропущено'
    }

    # TraceEvent кладём РЯДОМ со своей сборкой. Ссылка на неё строгоименная, то есть
    # связывается по ТОЧНОЙ версии: полагаться на чужую подпапку нельзя - на целевой базе
    # рядом может оказаться другая версия (EtwPerformanceProfiler несёт свою, 1.0.11),
    # и тогда наш приёмник не загрузится вовсе. Чужие папки при этом НЕ трогаем:
    # подмена чужой TraceEvent сломала бы чужой работающий профайлер.
    if ($copied) {
        $teTarget = Join-Path $addInDir (Split-Path $traceRef -Leaf)
        $teVer = [Diagnostics.FileVersionInfo]::GetVersionInfo($traceRef).FileVersion
        if ((Test-Path $teTarget) -and
            ([Diagnostics.FileVersionInfo]::GetVersionInfo($teTarget).FileVersion -eq $teVer)) {
            Step "TraceEvent $teVer уже на месте"
        } else {
            try { Copy-Item $traceRef $teTarget -Force; Step "выложено: TraceEvent $teVer" }
            catch { Fail "TraceEvent скопировать не удалось: $($_.Exception.Message)" }
        }
    }

    if (-not $copied -and $hadOld) {
        Copy-Item $backup $target -Force
        Step 'откат: возвращена прежняя сборка'
    }

    if ($svcWasRunning) {
        $startedAt = Get-Date
        Start-Service $service
        (Get-Service $service).WaitForStatus('Running', '00:03:00')
        if ((Get-Service $service).Status -ne 'Running' -and $hadOld) {
            Copy-Item $backup $target -Force
            Start-Service $service
            Fail 'служба не поднялась с новой сборкой, выполнен откат'
        } else {
            Step 'служба запущена, жду готовности к обслуживанию базы'
            if (Wait-NstReady $startedAt) { Step 'NST готов' }
            else { Fail 'NST не отчитался о готовности; компиляция C/AL пропущена'; $SkipCal = $true }
        }
    }
}

if ($SkipCal) {
    if ($errors.Count -gt 0) { exit 1 }
    Write-Host 'Готово (объекты не трогали).' -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------- объекты C/AL

Write-Host 'Объекты C/AL' -ForegroundColor Cyan

$objects = Select-String -Path $taskTxt -Pattern '^OBJECT\s+(\w+)\s+(\d+)\s' -Encoding utf8 |
    ForEach-Object { [pscustomobject]@{ Type = $_.Matches[0].Groups[1].Value; Id = [int]$_.Matches[0].Groups[2].Value } }
if (-not $objects) { throw "в $taskTxt не найдено ни одного OBJECT" }
Step ("в файле: " + (($objects | ForEach-Object { "$($_.Type) $($_.Id)" }) -join ', '))

& pwsh -NoProfile -File (Join-Path $PSScriptRoot 'convert-for-cside.ps1') $taskTxt | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'конвертация в cp866 не удалась' }
Step 'конвертировано в cp866'

function Invoke-Finsql([string]$argLine, [string]$logName) {
    $log = Join-Path $logDir $logName
    if ([IO.File]::Exists($log)) { [IO.File]::Delete($log) }
    $full = "$argLine,ServerName=$Server,Database=$Database,Ntauthentication=1,LogFile=`"$log`""
    Start-Process -FilePath $finsql -ArgumentList $full -Wait -NoNewWindow | Out-Null
    if ([IO.File]::Exists($log)) { return (Get-Content $log -Encoding Oem -Raw) }
    return ''
}

$impLog = Invoke-Finsql "Command=ImportObjects,File=`"$taskCp866`",ImportAction=Overwrite,SynchronizeSchemaChanges=Force" 'import.log'
if ($impLog) { Fail "импорт: $impLog" } else { Step 'импорт выполнен' }

# Кэш держит запущенный C/SIDE, в том числе зависший от прошлого прогона: без этого
# удаление падает с "Access to the path denied", а компиляция пойдёт против старой сборки.
Get-Process finsql -ErrorAction SilentlyContinue | ForEach-Object {
    Step "закрываю finsql (pid $($_.Id)) - он держит кэш сборки"
    Stop-Process -Id $_.Id -Force -Confirm:$false
}
$cache = Join-Path $env:TEMP "Microsoft Dynamics NAV\Add-Ins\$platform\LineProfiler_1.0.0.0"
if ([IO.Directory]::Exists($cache)) {
    try { [IO.Directory]::Delete($cache, $true); Step 'кэш сборки C/SIDE очищен' }
    catch { Fail "кэш сборки не удалось очистить ($($_.Exception.Message)); компиляция пойдёт против старой сборки" }
}

# Импорт сбросил Compiled в 0 у ВСЕХ импортированных объектов - компилируем все,
# в порядке зависимостей, и повторяем проход, если что-то не село с первого раза.
foreach ($pass in 1..2) {
    $pending = $objects | Where-Object { -not (Test-Compiled $_ $Server $Database) }
    if (-not $pending) { break }
    foreach ($o in ($pending | Sort-Object { $typeRank[$_.Type] })) {
        $log = Invoke-Finsql "Command=CompileObjects,Filter=`"Type=$($o.Type);ID=$($o.Id)`",SynchronizeSchemaChanges=Force" "compile-$($o.Type)-$($o.Id)-p$pass.log"
        if ($log) { Fail "компиляция $($o.Type) $($o.Id): $log" }
    }
}

# ---------------------------------------------------------------- проверка по базе

Write-Host 'Проверка' -ForegroundColor Cyan
$bad = @()
foreach ($o in $objects) {
    $c = Get-CompiledFlag $o $Server $Database
    if ($null -eq $c) { $bad += "$($o.Type) $($o.Id): нет строки в dbo.[Object]" }
    elseif ($c -ne 1) { $bad += "$($o.Type) $($o.Id): Compiled=$c" }
    else { Step "$($o.Type) $($o.Id): Compiled=1" }
}
foreach ($b in $bad) { Fail $b }

if ($errors.Count -gt 0) {
    Write-Host "ИТОГ: ошибок $($errors.Count)." -ForegroundColor Red
    exit 1
}
Write-Host 'ИТОГ: выложено и скомпилировано.' -ForegroundColor Green
exit 0
