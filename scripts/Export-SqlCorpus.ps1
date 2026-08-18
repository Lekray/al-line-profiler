#Requires -Version 5.1
# Выгружает корпус SQL из событий Application/705 в TSV (одна строка на запрос).
[CmdletBinding()]
param(
    [int]    $MaxEvents = 20000,
    [string] $OutFile
)
$ErrorActionPreference = 'Stop'
# По умолчанию пишем в out\: там боевые запросы не попадут под git add.
# Каталог scripts\ отслеживается, и корпус запросов рабочей базы туда класть нельзя.
if (-not $OutFile) { $OutFile = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'out') 'sql-corpus.tsv' }
$outDir = Split-Path $OutFile -Parent
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$events = @(Get-WinEvent -FilterHashtable @{LogName='Application'; Id=705} -MaxEvents $MaxEvents -ErrorAction SilentlyContinue)
$rows = New-Object System.Collections.Generic.List[string]
$rows.Add((@('Time','ExecMs','Category','Kind','Sql') -join "`t"))

foreach ($ev in $events) {
    $msg = $ev.Message
    if (-not $msg) { continue }
    $lines = $msg -split "`r`n|`n"
    $cat = ''; $ms = ''; $kind = ''
    $sqlStart = -1; $sqlEnd = -1
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $l = $lines[$i]
        if ($l -match '^Category:\s*(.+)$')            { $cat  = $Matches[1].Trim() }
        elseif ($l -match '^\s*Execution time:\s*(\d+)') { $ms  = $Matches[1] }
        elseif ($l -match '^\s*Message:\s*(.+)$')      { $kind = $Matches[1].Trim() }
        elseif ($l -match '^\s*Connection ID:')        { $sqlStart = $i + 1 }
        elseif ($l -match '^ProcessId:')               { $sqlEnd = $i - 1; break }
    }
    if ($sqlStart -lt 0) { continue }
    if ($sqlEnd -lt $sqlStart) { $sqlEnd = $lines.Length - 1 }
    $sql = ($lines[$sqlStart..$sqlEnd] -join ' ').Trim()
    $sql = $sql -replace '\s+', ' '
    if ($sql.Length -eq 0) { continue }
    $rows.Add((@($ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'), $ms, $cat, $kind, $sql) -join "`t"))
}

$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutFile, (($rows -join "`r`n") + "`r`n"), $enc)
Write-Host ("Событий: {0}, запросов: {1} -> {2}" -f $events.Count, ($rows.Count - 1), $OutFile)
