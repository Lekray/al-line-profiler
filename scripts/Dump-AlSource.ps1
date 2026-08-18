<#
.SYNOPSIS
    Выгружает исходный код C/AL объектов базы NAV из таблицы Object Metadata.

.DESCRIPTION
    Поле [Object Metadata].[User AL Code] хранит исходник скомпилированного объекта:
    4 байта заголовка (02 45 7D 5B), далее raw Deflate, внутри UTF-8 с переносами CRLF.
    Это тот самый массив строк, по которому платформа нумерует строки кода — и в событиях
    трассировки, и во встроенном Code Coverage.

    ВНИМАНИЕ: нумерация строк здесь НЕ совпадает с нумерацией текстового экспорта C/SIDE
    (файлы baseline). Экспорт содержит разметку объекта и заметно длиннее. Источником
    номеров строк для профайлера может быть только этот дамп.

    Скрипт ничего не пишет в базу: только SELECT. Зависимостей нет — System.Data.SqlClient
    и System.IO.Compression входят в состав .NET Framework.

.PARAMETER Server
    Экземпляр SQL Server. По умолчанию localhost.

.PARAMETER Database
    База данных NAV. По умолчанию NAV либо значение LP_DATABASE.

.PARAMETER OutDir
    Куда складывать. По умолчанию <корень репозитория>\.alsrc (вне git).

.PARAMETER ObjectType
    Фильтр по типам объектов: 1 Table, 2 Form, 3 Report, 4 Dataport, 5 Codeunit,
    6 XMLport, 7 MenuSuite, 8 Page, 9 Query, 10 System. По умолчанию — все.

.PARAMETER FromId, ToId
    Фильтр по диапазону номеров объектов.

.EXAMPLE
    .\Dump-AlSource.ps1
    Выгружает все объекты базы в .alsrc\.

.EXAMPLE
    .\Dump-AlSource.ps1 -ObjectType 5 -FromId 80 -ToId 80
    Выгружает один кодюнит.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string] $Server   = 'localhost',
    [string] $Database = $(if ($env:LP_DATABASE) { $env:LP_DATABASE } else { 'NAV' }),
    [string] $OutDir,
    [int[]]  $ObjectType,
    [int]    $FromId = 0,
    [int]    $ToId   = 2147483647,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'

# --- пути -------------------------------------------------------------------
if (-not $OutDir) {
    # scripts -> корень репозитория
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $OutDir   = Join-Path $repoRoot '.alsrc'
}
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

# --- константы --------------------------------------------------------------
$MAGIC = [byte[]] @(0x02, 0x45, 0x7D, 0x5B)
$TYPE_NAMES = @{
    1 = 'Table'; 2 = 'Form';      3 = 'Report'; 4 = 'Dataport'; 5 = 'Codeunit'
    6 = 'XMLport'; 7 = 'MenuSuite'; 8 = 'Page'; 9 = 'Query';   10 = 'System'
}

# --- запрос -----------------------------------------------------------------
$where = @('m.[User AL Code] IS NOT NULL', 'm.[Object ID] BETWEEN @from AND @to')
if ($ObjectType) {
    $list = ($ObjectType | ForEach-Object { [int]$_ }) -join ','
    $where += "m.[Object Type] IN ($list)"
}
$sql = @"
SELECT  m.[Object Type]            AS ObjType,
        m.[Object ID]              AS ObjId,
        ISNULL(o.[Name], '')       AS ObjName,
        ISNULL(o.[Compiled], 0)    AS Compiled,
        o.[Date]                   AS ObjDate,
        o.[Time]                   AS ObjTime,
        ISNULL(o.[Version List],'') AS VersionList,
        ISNULL(m.[Hash], '')       AS MetaHash,
        CONVERT(varbinary(max), m.[User AL Code]) AS Src
FROM    dbo.[Object Metadata] m
LEFT JOIN dbo.[Object] o
       ON  o.[Type] = m.[Object Type]
       AND o.[ID]   = m.[Object ID]
       AND o.[Company Name] = ''
WHERE   $($where -join ' AND ')
ORDER BY m.[Object Type], m.[Object ID]
"@

# --- выгрузка ---------------------------------------------------------------
$sw    = [System.Diagnostics.Stopwatch]::StartNew()
$rows  = New-Object System.Collections.Generic.List[string]
$ok = 0; $failed = 0; $badMagic = 0; $totalLines = 0; $totalBytes = 0
$failedList = New-Object System.Collections.Generic.List[string]

$connStr = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True"
$cn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
    $cn.Open()
    $cmd = $cn.CreateCommand()
    $cmd.CommandText    = $sql
    $cmd.CommandTimeout = 600
    [void]$cmd.Parameters.AddWithValue('@from', $FromId)
    [void]$cmd.Parameters.AddWithValue('@to',   $ToId)

    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $t    = [int]$reader['ObjType']
        $id   = [int]$reader['ObjId']
        $name = [string]$reader['ObjName']
        $blob = [byte[]]$reader['Src']

        # заголовок обязателен: без него это не тот формат, который мы умеем читать
        if ($blob.Length -le 4 -or
            $blob[0] -ne $MAGIC[0] -or $blob[1] -ne $MAGIC[1] -or
            $blob[2] -ne $MAGIC[2] -or $blob[3] -ne $MAGIC[3]) {
            $badMagic++
            $failedList.Add(("{0} {1} {2} — чужой заголовок" -f $TYPE_NAMES[$t], $id, $name))
            continue
        }

        try {
            $inMs  = New-Object System.IO.MemoryStream (,$blob[4..($blob.Length - 1)])
            $defl  = New-Object System.IO.Compression.DeflateStream($inMs, [System.IO.Compression.CompressionMode]::Decompress)
            $outMs = New-Object System.IO.MemoryStream
            $defl.CopyTo($outMs)
            $bytes = $outMs.ToArray()
            $defl.Dispose(); $inMs.Dispose(); $outMs.Dispose()
        }
        catch {
            $failed++
            $failedList.Add(("{0} {1} {2} — {3}" -f $TYPE_NAMES[$t], $id, $name, $_.Exception.Message))
            continue
        }

        $path = Join-Path $OutDir ("{0}_{1}.al" -f $t, $id)
        [System.IO.File]::WriteAllBytes($path, $bytes)

        $text  = [System.Text.Encoding]::UTF8.GetString($bytes)
        $lines = ($text -split "`r`n").Count

        $dt = ''
        if ($reader['ObjDate'] -isnot [DBNull]) { $dt = ([datetime]$reader['ObjDate']).ToString('yyyy-MM-dd') }
        $tm = ''
        if ($reader['ObjTime'] -isnot [DBNull]) { $tm = ([datetime]$reader['ObjTime']).ToString('HH:mm:ss') }

        $rows.Add((@(
            $t
            $TYPE_NAMES[$t]
            $id
            $name
            $lines
            $bytes.Length
            [int]$reader['Compiled']
            $dt
            $tm
            ([string]$reader['VersionList'])
            ([string]$reader['MetaHash'])
        ) -join "`t"))

        $ok++; $totalLines += $lines; $totalBytes += $bytes.Length
        if (-not $Quiet -and ($ok % 1000) -eq 0) {
            Write-Host ("  выгружено {0}..." -f $ok) -ForegroundColor DarkGray
        }
    }
    $reader.Close()
}
finally {
    if ($cn.State -ne 'Closed') { $cn.Close() }
}

# --- индекс -----------------------------------------------------------------
$header = @('Type','TypeName','ID','Name','Lines','Bytes','Compiled','Date','Time','VersionList','Hash') -join "`t"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText(
    (Join-Path $OutDir 'index.tsv'),
    (($header + "`r`n") + (($rows -join "`r`n") + "`r`n")),
    $utf8NoBom)

$sw.Stop()

# --- итог (компактно, под один экран) ---------------------------------------
Write-Host ''
Write-Host ("База:      {0}\{1}" -f $Server, $Database)
Write-Host ("Каталог:   {0}" -f $OutDir)
Write-Host ("Объектов:  {0}" -f $ok)
Write-Host ("Строк:     {0:N0}" -f $totalLines)
Write-Host ("Объём:     {0:N1} МБ" -f ($totalBytes / 1MB))
Write-Host ("Время:     {0:N1} с" -f $sw.Elapsed.TotalSeconds)
if ($badMagic -gt 0 -or $failed -gt 0) {
    Write-Host ("Пропущено: {0} (заголовок {1}, распаковка {2})" -f ($badMagic + $failed), $badMagic, $failed) -ForegroundColor Yellow
    $failedList | Select-Object -First 10 | ForEach-Object { Write-Host ("  {0}" -f $_) -ForegroundColor DarkYellow }
} else {
    Write-Host 'Пропущено: 0' -ForegroundColor Green
}
