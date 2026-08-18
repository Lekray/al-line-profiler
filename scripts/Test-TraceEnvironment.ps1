<#
.SYNOPSIS
    Проверяет, готова ли машина к сбору трассировки C/AL скриптом Collect-AlTrace.ps1.

.DESCRIPTION
    Двенадцать проверок одной таблицей: версия платформы и PowerShell, наличие
    провайдера ETW и события уровня оператора с именами полей payload, канал
    журнала и его размер, ключ EnableFullALFunctionTracing в CustomSettings.config,
    состояние службы, доступность SQL и права запустившего.

    Событие уровня оператора ищется НЕ по номеру: в манифесте провайдера
    отбирается событие, в payload которого есть поля lineNumber и statement.
    Между версиями платформы номер события и порядок полей меняются, а имена нет -
    готовый профайлер сообщества сломался именно на жёстко зашитых индексах.

    Скрипт только читает: настройки не меняет, службу не трогает, в базу шлёт
    единственный SELECT COUNT(*). Прав администратора не требует - но сообщает,
    есть ли они, потому что самому сбору они нужны.

.PARAMETER ServerInstance
    Имя инстанции NAV. По умолчанию DynamicsNAV110.

.PARAMETER Server
    Экземпляр SQL Server. По умолчанию localhost.

.PARAMETER Database
    База данных NAV. По умолчанию NAV либо значение LP_DATABASE.

.PARAMETER NavServicePath
    Каталог Service установки NAV. По умолчанию берётся из ImagePath службы, а если
    службы нет - ищется в C:\Program Files\Microsoft Dynamics NAV\*\Service.

.EXAMPLE
    .\Test-TraceEnvironment.ps1
    Проверяет инстанцию DynamicsNAV110 и базу по умолчанию на localhost.

.EXAMPLE
    .\Test-TraceEnvironment.ps1 -ServerInstance DynamicsNAV110 -Database CRONUS
    То же, но с другой базой.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [string] $ServerInstance = 'DynamicsNAV110',
    [string] $Server         = 'localhost',
    [string] $Database       = $(if ($env:LP_DATABASE) { $env:LP_DATABASE } else { 'NAV' }),
    [string] $NavServicePath
)

$ErrorActionPreference = 'Stop'

$PROVIDER    = 'Microsoft-DynamicsNAV-Server'
$SETTING_KEY = 'EnableFullALFunctionTracing'

# --- таблица результатов ----------------------------------------------------
# ширины подобраны так, чтобы имена полей payload и путь к установке NAV
# укладывались целиком, а вся таблица влезала в один скриншот
$W_CHECK = 28
$W_VALUE = 64

$rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param([string] $Check, [string] $Value, [string] $Verdict)
    if ($null -eq $Value) { $Value = '' }
    if ($Value.Length -gt $W_VALUE) { $Value = $Value.Substring(0, $W_VALUE - 3) + '...' }
    [void]$rows.Add([pscustomobject]@{ Check = $Check; Value = $Value; Verdict = $Verdict })
}

# --- 1. PowerShell ----------------------------------------------------------
$psv = '{0} ({1})' -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition
if ($PSVersionTable.PSEdition -eq 'Core') { Add-Row 'Windows PowerShell' $psv 'НЕТ' }
else                                      { Add-Row 'Windows PowerShell' $psv 'ОК'  }

# --- 2. права ---------------------------------------------------------------
$wi   = [Security.Principal.WindowsIdentity]::GetCurrent()
$wp   = New-Object Security.Principal.WindowsPrincipal($wi)
$elev = $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($elev) { Add-Row 'Права администратора' 'есть' 'ОК' }
else       { Add-Row 'Права администратора' 'нет - сбор запускать из поднятой консоли' 'ВНИМ' }

# --- 3. платформа NAV -------------------------------------------------------
$svcRoot = $null
if ($NavServicePath) {
    if (Test-Path -LiteralPath $NavServicePath) { $svcRoot = (Resolve-Path -LiteralPath $NavServicePath).Path }
} else {
    $key = 'HKLM:\SYSTEM\CurrentControlSet\Services\MicrosoftDynamicsNavServer$' + $ServerInstance
    $img = $null
    try { $img = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).ImagePath } catch { }
    if ($img -and $img -match '^"([^"]+)"') {
        $d = Split-Path -Parent $Matches[1]
        if (Test-Path -LiteralPath $d) { $svcRoot = $d }
    }
    if (-not $svcRoot) {
        $probe = @(Get-ChildItem 'C:\Program Files\Microsoft Dynamics NAV' -Directory -ErrorAction SilentlyContinue |
                   ForEach-Object { Join-Path $_.FullName 'Service' } |
                   Where-Object { Test-Path -LiteralPath $_ } | Sort-Object -Descending)
        if ($probe.Count -gt 0) { $svcRoot = $probe[0] }
    }
}
if ($svcRoot) {
    $exe = Join-Path $svcRoot 'Microsoft.Dynamics.Nav.Server.exe'
    if (Test-Path -LiteralPath $exe) {
        Add-Row 'Платформа NAV' ((Get-Item -LiteralPath $exe).VersionInfo.FileVersion + ' ' + $svcRoot) 'ОК'
    } else {
        Add-Row 'Платформа NAV' ('нет Server.exe в ' + $svcRoot) 'НЕТ'
    }
} else {
    Add-Row 'Платформа NAV' 'установка не найдена, укажите -NavServicePath' 'НЕТ'
}

# --- 4. служба инстанции ----------------------------------------------------
$svc = Get-Service -Name ('MicrosoftDynamicsNavServer$' + $ServerInstance) -ErrorAction SilentlyContinue
if (-not $svc)                       { Add-Row 'Служба NST' ('нет службы инстанции ' + $ServerInstance) 'НЕТ' }
elseif ($svc.Status -eq 'Running')   { Add-Row 'Служба NST' ('{0} / {1}' -f $svc.Status, $svc.StartType) 'ОК' }
else                                 { Add-Row 'Служба NST' ('{0} / {1} - запустить перед сбором' -f $svc.Status, $svc.StartType) 'ВНИМ' }

# --- 5..7. провайдер ETW, событие оператора, поля payload -------------------
$prov = $null
try { $prov = Get-WinEvent -ListProvider $PROVIDER -ErrorAction Stop } catch { }

$stmt = $null
$channel = $PROVIDER + '/Debug'
if (-not $prov) {
    Add-Row 'Провайдер ETW' ($PROVIDER + ' не зарегистрирован') 'НЕТ'
    Add-Row 'Событие оператора' 'нет провайдера' 'НЕТ'
    Add-Row 'Поля payload' '-' 'НЕТ'
} else {
    Add-Row 'Провайдер ETW' ('{0}, событий {1}' -f $PROVIDER, $prov.Events.Count) 'ОК'

    # ищем по именам полей, а не по номеру события
    $rxField = New-Object System.Text.RegularExpressions.Regex('<data\s+name\s*=\s*"([^"]+)"', 'IgnoreCase')
    foreach ($e in $prov.Events) {
        if (-not $e.Template) { continue }
        $names = @()
        foreach ($m in $rxField.Matches($e.Template)) { $names += $m.Groups[1].Value }
        if (($names -contains 'lineNumber') -and ($names -contains 'statement')) {
            if ($null -eq $stmt -or [int]$e.Version -gt $stmt.Version) {
                $task = ''
                if ($e.Task) { $task = [string]$e.Task.Name }
                $ch = ''
                if ($e.LogLink) { $ch = [string]$e.LogLink.LogName }
                $stmt = [pscustomobject]@{
                    Id = [int]$e.Id; Version = [int]$e.Version; Task = $task; Channel = $ch; Fields = $names
                }
            }
        }
    }
    if ($stmt) {
        # канал берётся из манифеста: в другой версии платформы он может быть иным
        if ($stmt.Channel) { $channel = $stmt.Channel }
        Add-Row 'Событие оператора' ('Id={0} v{1} {2}, полей {3}' -f
            $stmt.Id, $stmt.Version, $stmt.Task, $stmt.Fields.Count) 'ОК'
        $key6 = @($stmt.Fields | Where-Object {
            @('sessionId','objectType','objectId','functionName','lineNumber','statement') -contains $_ })
        Add-Row 'Поля payload' ($key6 -join ',') 'ОК'
    } else {
        Add-Row 'Событие оператора' 'нет события с полями lineNumber и statement' 'НЕТ'
        Add-Row 'Поля payload' '-' 'НЕТ'
    }
}

# --- 8..9. канал журнала ----------------------------------------------------
# wevtutil читается без прав администратора; keywords и level в выводе появляются
# только когда фильтр задан - их отсутствие означает "без фильтра"
$prevEap = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
$gl = & wevtutil.exe gl $channel 2>&1 | ForEach-Object { [string]$_ }
$glCode = $LASTEXITCODE
$ErrorActionPreference = $prevEap

if ($glCode -ne 0) {
    Add-Row 'Канал журнала' ($channel + ' не зарегистрирован') 'НЕТ'
    Add-Row 'Размер канала' '-' 'НЕТ'
} else {
    $chEnabled = $false; $chMax = 0L; $chRet = $false
    foreach ($line in $gl) {
        $t = $line.Trim()
        if     ($t -match '^enabled:\s*(\S+)')   { $chEnabled = ($Matches[1] -eq 'true') }
        elseif ($t -match '^maxSize:\s*(\d+)')   { $chMax     = [int64]$Matches[1] }
        elseif ($t -match '^retention:\s*(\S+)') { $chRet     = ($Matches[1] -eq 'true') }
    }
    if ($chEnabled) { $st = 'включён (сбор его выключит и очистит)' } else { $st = 'выключен' }
    Add-Row 'Канал журнала' ('{0} - {1}' -f $channel, $st) 'ОК'

    $sz = '{0:N1} МБ, retention={1}' -f ($chMax / 1MB), $chRet.ToString().ToLower()
    if ($chMax -lt (64MB)) {
        # заполнившись, канал с retention=true молча перестаёт писать
        Add-Row 'Размер канала' ($sz + ' - мало, сбор поднимет') 'ВНИМ'
    } else {
        Add-Row 'Размер канала' $sz 'ОК'
    }
}

# --- 10..11. CustomSettings.config ------------------------------------------
$cfgPath = $null
if ($svcRoot) {
    $c = Join-Path $svcRoot ('Instances\{0}\CustomSettings.config' -f $ServerInstance)
    if (Test-Path -LiteralPath $c) { $cfgPath = $c }
    else {
        $c = Join-Path $svcRoot 'CustomSettings.config'
        if (Test-Path -LiteralPath $c) { $cfgPath = $c }
    }
}
if (-not $cfgPath) {
    Add-Row 'CustomSettings.config' 'не найден' 'НЕТ'
    Add-Row $SETTING_KEY '-' 'НЕТ'
} else {
    # путь показываем относительно каталога Service - он уже есть в строке "Платформа NAV"
    $shown = $cfgPath
    if ($svcRoot -and $cfgPath.StartsWith($svcRoot, [StringComparison]::OrdinalIgnoreCase)) {
        $shown = $cfgPath.Substring($svcRoot.Length).TrimStart('\')
    }
    $canWrite = $false
    try { $fs = [System.IO.File]::Open($cfgPath, 'Open', 'Write', 'None'); $fs.Close(); $canWrite = $true } catch { }
    if ($canWrite) { Add-Row 'CustomSettings.config' ($shown + ', запись доступна') 'ОК' }
    else           { Add-Row 'CustomSettings.config' ($shown + ', только чтение - сбор не сможет включить трассировку') 'ВНИМ' }

    $txt = [System.IO.File]::ReadAllText($cfgPath, [System.Text.Encoding]::UTF8)
    $m = [regex]::Match($txt, ('<add\s+key\s*=\s*"{0}"\s+value\s*=\s*"([^"]*)"' -f $SETTING_KEY),
                        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) {
        Add-Row $SETTING_KEY 'ключа нет - сбор его добавит, нужен перезапуск NST' 'ВНИМ'
    } elseif ($m.Groups[1].Value -eq 'true') {
        Add-Row $SETTING_KEY 'true - трассировка уже включена' 'ОК'
    } else {
        Add-Row $SETTING_KEY ($m.Groups[1].Value + ' - сбор включит, нужен перезапуск NST') 'ВНИМ'
    }
}

# --- 12. SQL ----------------------------------------------------------------
try {
    $cs = 'Server={0};Database={1};Integrated Security=True;TrustServerCertificate=True;Connect Timeout=5' -f $Server, $Database
    $cn = New-Object System.Data.SqlClient.SqlConnection $cs
    $cn.Open()
    try {
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = 'SELECT COUNT(*) FROM dbo.[Object Metadata] WHERE [User AL Code] IS NOT NULL'
        $n = [int]$cmd.ExecuteScalar()
        Add-Row 'SQL, исходники C/AL' ('{0}\{1}: объектов с кодом {2:N0}' -f $Server, $Database, $n) 'ОК'
    }
    finally { $cn.Close() }
}
catch {
    Add-Row 'SQL, исходники C/AL' ('{0}\{1}: {2}' -f $Server, $Database, $_.Exception.Message) 'НЕТ'
}

# ---------------------------------------------------------------------------
# вывод (компактно, под один скриншот)
# ---------------------------------------------------------------------------
$sep = ('-' * $W_CHECK) + '  ' + ('-' * $W_VALUE) + '  -------'
Write-Host ''
Write-Host (('{0,-' + $W_CHECK + '}  {1,-' + $W_VALUE + '}  {2}') -f 'ПРОВЕРКА', 'ЗНАЧЕНИЕ', 'ВЕРДИКТ')
Write-Host $sep
foreach ($r in $rows) {
    $color = 'Green'
    if     ($r.Verdict -eq 'НЕТ')  { $color = 'Red' }
    elseif ($r.Verdict -eq 'ВНИМ') { $color = 'Yellow' }
    Write-Host ((('{0,-' + $W_CHECK + '}  {1,-' + $W_VALUE + '}  ') -f $r.Check, $r.Value)) -NoNewline
    Write-Host $r.Verdict -ForegroundColor $color
}
Write-Host $sep

$bad  = @($rows | Where-Object { $_.Verdict -eq 'НЕТ'  }).Count
$warn = @($rows | Where-Object { $_.Verdict -eq 'ВНИМ' }).Count
if ($bad -gt 0) {
    Write-Host ('Готовность: НЕТ - блокирующих {0}, предупреждений {1}' -f $bad, $warn) -ForegroundColor Red
} elseif ($warn -gt 0) {
    Write-Host ('Готовность: ДА с оговорками - предупреждений {0}' -f $warn) -ForegroundColor Yellow
} else {
    Write-Host 'Готовность: ДА' -ForegroundColor Green
}
if ($stmt) {
    Write-Host ('Полный payload события {0}: {1}' -f $stmt.Id, ($stmt.Fields -join ',')) -ForegroundColor DarkGray
}

if ($bad -gt 0) { exit 1 }
exit 0
