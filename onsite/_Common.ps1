# Общее для всех шагов пакета. Отдельным файлом, потому что каждый шаг запускают
# по отдельности и в свой день.
#
# ВАЖНО: пакет рассчитан на Windows PowerShell 5.1 - на сервере NAV 2017 pwsh 7
# может не оказаться вовсе. Никаких операторов ?? ?. и тернарных, никакого -Parallel.

$ErrorActionPreference = 'Stop'

function Line([string]$Label, $Value) {
    # Ровная колонка: снимок экрана читают глазами, а не разбирают программой.
    Write-Host ($Label.PadRight(22) + ': ') -NoNewline
    Write-Host $Value
}

function Head([string]$Marker) {
    Write-Host ''
    Write-Host "=== $Marker ===" -ForegroundColor Cyan
}

function Bad([string]$Text)  { Write-Host "  !! $Text" -ForegroundColor Red }
function Warn([string]$Text) { Write-Host "  ?  $Text" -ForegroundColor Yellow }
function Good([string]$Text) { Write-Host "  OK $Text" -ForegroundColor Green }

function Get-NavContext {
    param([string]$Instance)

    $ctx = @{}
    $filter = "Name LIKE 'MicrosoftDynamicsNavServer%'"
    $svcs = @(Get-CimInstance Win32_Service -Filter $filter)
    if ($svcs.Count -eq 0) { throw 'служба NAV не найдена на этой машине' }
    if ($Instance) {
        $svcs = @($svcs | Where-Object { $_.Name -like "*`$$Instance" })
        if ($svcs.Count -eq 0) { throw "инстанция '$Instance' не найдена" }
    }
    $svc = $svcs[0]

    $ctx.ServiceName = $svc.Name
    $ctx.Instance    = ($svc.Name -split '\$')[-1]
    $ctx.Running     = ($svc.State -eq 'Running')
    $ctx.Account     = $svc.StartName
    $ctx.AllServices = ($svcs | ForEach-Object { ($_.Name -split '\$')[-1] }) -join ', '

    # Корень NAV берём из пути службы: зашивать 110/100 нельзя, номер платформы разный.
    $exe = $svc.PathName
    if ($exe -match '^"([^"]+)"') { $exe = $Matches[1] } else { $exe = ($exe -split ' ')[0] }
    $ctx.ServiceExe = $exe
    $ctx.Root       = Split-Path $exe -Parent
    $ctx.NavHome    = Split-Path $ctx.Root -Parent
    $ctx.PlatformNo = Split-Path $ctx.NavHome -Leaf

    if (Test-Path $exe) {
        $ctx.Version = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe).FileVersion
    } else {
        $ctx.Version = '(файл службы не найден)'
    }

    # Конфигурация инстанции: в NAV 2016+ лежит в Instances\<имя>, в более старых - в корне.
    $cfg = Join-Path $ctx.Root ('Instances\' + $ctx.Instance + '\CustomSettings.config')
    if (-not (Test-Path $cfg)) { $cfg = Join-Path $ctx.Root 'CustomSettings.config' }
    $ctx.ConfigPath = $cfg
    $ctx.Settings = @{}
    if (Test-Path $cfg) {
        # Корневой элемент бывает и <configuration><appSettings>, и просто <appSettings>:
        # на стенде 2018 он второй. Берём узлы поиском, а не по пути.
        [xml]$x = Get-Content $cfg
        foreach ($a in $x.SelectNodes('//add')) { if ($a.key) { $ctx.Settings[$a.key] = $a.value } }
    }
    $ctx.DbServer = $ctx.Settings['DatabaseServer']
    if ($ctx.Settings['DatabaseInstance']) { $ctx.DbServer = $ctx.DbServer + '\' + $ctx.Settings['DatabaseInstance'] }
    $ctx.DbName   = $ctx.Settings['DatabaseName']
    $ctx.Company  = $ctx.Settings['ServicesDefaultCompany']

    # finsql нужен для импорта объектов; он ставится с клиентом, рядом или в x86.
    $ctx.Finsql = ''
    $roots = @($ctx.NavHome)
    $roots += ($ctx.NavHome -replace 'Program Files \(x86\)', 'Program Files')
    $roots += ($ctx.NavHome -replace 'Program Files(?! \(x86\))', 'Program Files (x86)')
    foreach ($r in ($roots | Select-Object -Unique)) {
        $c = Join-Path $r 'RoleTailored Client\finsql.exe'
        if (Test-Path $c) { $ctx.Finsql = $c; break }
    }

    $ctx.Csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
    $ctx.AddIns = Join-Path $ctx.Root 'Add-ins'
    return $ctx
}

function Get-TraceEventDlls {
    param($Ctx)
    # Поставочный приёмник NAV лежит в своей подпапке, а на целевой базе может быть
    # только чужой, принесённый сторонним профайлером. Годится любой - берём старший.
    if (-not (Test-Path $Ctx.AddIns)) { return @() }
    $found = Get-ChildItem $Ctx.AddIns -Recurse -Filter 'Microsoft.Diagnostics.Tracing.TraceEvent.dll' -ErrorAction SilentlyContinue
    $list = @()
    foreach ($f in $found) {
        $list += [pscustomobject]@{
            Version = [Diagnostics.FileVersionInfo]::GetVersionInfo($f.FullName).FileVersion
            Path    = $f.FullName
        }
    }
    return ($list | Sort-Object { [version]$_.Version } -Descending)
}

function Invoke-NavSql {
    param($Ctx, [string]$Sql)
    Add-Type -AssemblyName System.Data
    $cs = "Server=$($Ctx.DbServer);Database=$($Ctx.DbName);Integrated Security=True;TrustServerCertificate=True"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 120
        $rows = @()
        $r = $cmd.ExecuteReader()
        while ($r.Read()) {
            $o = @{}
            for ($i = 0; $i -lt $r.FieldCount; $i++) { $o[$r.GetName($i)] = $r.GetValue($i) }
            $rows += [pscustomobject]$o
        }
        $r.Close()
        return $rows
    } finally { $conn.Close() }
}

function Invoke-NavSqlNonQuery {
    param($Ctx, [string]$Sql)
    Add-Type -AssemblyName System.Data
    $cs = "Server=$($Ctx.DbServer);Database=$($Ctx.DbName);Integrated Security=True;TrustServerCertificate=True"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 120
        return $cmd.ExecuteNonQuery()
    } finally { $conn.Close() }
}
