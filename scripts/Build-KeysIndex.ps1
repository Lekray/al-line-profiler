#Requires -Version 5.1
<#
.SYNOPSIS
    Собирает справочники по таблицам NAV из текстового экспорта C/SIDE: ключи, поля и
    словарь SQL-имён таблиц. Основа советника по ключам в LineProfiler.

.DESCRIPTION
    Источник — подпапка Tables каталога baseline: текстовый экспорт таблиц из C/SIDE
    (UTF-8/CRLF). Каталог берётся из -BaseDir, иначе из переменной окружения
    LP_BASELINE_DIR, иначе это папка baseline в корне репозитория.

    Готовый индекс метаданных baseline (tables.keys.tsv), если такой ведётся, для
    советника непригоден:
      - выключенные ключи (Enabled=No) в него не попадают, из-за чего номера уцелевших
        ключей съезжают; имя SQL-индекса платформа выводит именно из номера ключа,
        поэтому нумерация здесь сквозная по факту, вместе с выключенными. Проверено на
        dbo.[NAV$G_L Entry]: ключ N даёт индекс $(N-1), ключ 1 (кластерный) — индекс
        <SQL-имя таблицы>$0, а выключенные ключи 4 и 5 съедают номера $3 и $4, которых
        в sys.indexes нет. Со сдвинутыми номерами советник указал бы не на тот индекс;
      - Clustered теряется, когда запись ключа перенесена на следующую строку;
      - SumIndexFields, MaintainSQLIndex и MaintainSIFTIndex не извлекаются вовсе.
    Записи секций KEYS и FIELDS здесь сначала склеиваются по продолжениям строк и лишь
    потом разбираются; свойства читаются с отрицательным ретроспективным поиском, иначе
    SQLIndex= ловится внутри MaintainSQLIndex=.

    Результат — три TSV в <корень репозитория>\out (UTF-8 без BOM, CRLF):
      keys.tsv    TableID, TableName, KeyNo, Enabled, KeyFields, SumIndexFields,
                  Clustered, MaintainSQLIndex, MaintainSIFTIndex, SQLIndex
      fields.tsv  TableID, TableName, FieldNo, FieldName, DataType, FieldClass,
                  CalcFormula, Enabled
      sqlmap.tsv  SqlName, TableID, TableName

    Словарь SQL-имён строится ПРЯМЫМ преобразованием имени таблицы: обратное
    преобразование невозможно — подчёркивание неоднозначно. Платформа заменяет на
    подчёркивание символы . " \ / ' % ] [ ; пробел она НЕ заменяет — проверено на
    фактических именах рабочей базы (dbo.[NAV$G_L Entry], dbo.[Access Control]).
    Ключ -SpaceToUnderscore включает замену пробела, если она понадобится на другой среде.

    В выгрузке встречаются устаревшие файлы переименованных таблиц (два файла на один
    номер). Побеждает запись с более свежими Date/Time из OBJECT-PROPERTIES, остальные
    в справочники не попадают и считаются отдельно.

    Скрипт идемпотентен: перезапуск перезаписывает три файла. В базу ничего не пишется —
    только SELECT для сверки словаря с sys.tables (отключается ключом -SkipSqlCheck).
    Внешних зависимостей нет: System.Data.SqlClient входит в .NET Framework.

.PARAMETER BaseDir
    Каталог с экспортом таблиц: подпапка Tables каталога baseline. По умолчанию берётся
    из переменной окружения LP_BASELINE_DIR, а без неё — <корень репозитория>\baseline.

.PARAMETER OutDir
    Куда складывать TSV. По умолчанию <корень репозитория>\out.

.PARAMETER Server
    Экземпляр SQL Server для сверки словаря. По умолчанию localhost.

.PARAMETER Database
    База данных NAV для сверки словаря. По умолчанию NAV либо значение LP_DATABASE.

.PARAMETER SkipSqlCheck
    Не обращаться к SQL: собрать только справочники, сверку словаря пропустить.

.PARAMETER SpaceToUnderscore
    Заменять пробел на подчёркивание при построении SQL-имени. По умолчанию выключено:
    в NAV 2018 пробел в имени таблицы сохраняется.

.PARAMETER Quiet
    Не печатать прогресс.

.EXAMPLE
    .\Build-KeysIndex.ps1
    Разбирает весь каталог baseline и сверяет словарь имён с базой на localhost.

.EXAMPLE
    .\Build-KeysIndex.ps1 -SkipSqlCheck
    Только справочники, без обращения к SQL (годится на машине без доступа к базе).
#>
[CmdletBinding()]
param(
    [string] $BaseDir,
    [string] $OutDir,
    [string] $Server   = 'localhost',
    [string] $Database = $(if ($env:LP_DATABASE) { $env:LP_DATABASE } else { 'NAV' }),
    [switch] $SkipSqlCheck,
    [switch] $SpaceToUnderscore,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'

# --- пути -------------------------------------------------------------------
# scripts -> корень репозитория
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $BaseDir) {
    # baseline — текстовый экспорт объектов целевой базы; в репозиторий он не входит
    $baseRoot = $env:LP_BASELINE_DIR
    if (-not $baseRoot) { $baseRoot = Join-Path $repoRoot 'baseline' }
    $BaseDir = Join-Path $baseRoot 'Tables'
}
if (-not $OutDir)  { $OutDir  = Join-Path $repoRoot 'out' }

if (-not (Test-Path -LiteralPath $BaseDir)) {
    Write-Host ''
    Write-Host ("Не найден каталог экспорта таблиц: {0}" -f $BaseDir) -ForegroundColor Red
    Write-Host 'Baseline — текстовый экспорт объектов целевой базы из C/SIDE; в репозитории его нет.' -ForegroundColor Yellow
    Write-Host 'Выгрузите объекты (File -> Export) в каталог с подпапками Tables, Pages, Codeunits, ...,' -ForegroundColor Yellow
    Write-Host 'а путь к нему задайте переменной окружения LP_BASELINE_DIR или параметром -BaseDir.' -ForegroundColor Yellow
    exit 1
}
$files = @(Get-ChildItem -LiteralPath $BaseDir -Filter '*.txt' -File)
if ($files.Count -eq 0) {
    Write-Host ''
    Write-Host ("В каталоге нет файлов *.txt: {0}" -f $BaseDir) -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

# --- разбор экспорта --------------------------------------------------------

# Символы, которые платформа заменяет на подчёркивание в имени SQL-объекта.
$SQL_BAD = '."\/' + "'" + '%]['

function ConvertTo-SqlName {
    param([string] $Name)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($SQL_BAD.IndexOf($ch) -ge 0)             { [void]$sb.Append('_') }
        elseif ($SpaceToUnderscore -and $ch -eq ' ') { [void]$sb.Append('_') }
        else                                         { [void]$sb.Append($ch) }
    }
    $sb.ToString()
}

# Записи секции: начинаются со строки '    {', продолжения приклеиваются через пробел.
# Так восстанавливаются свойства, перенесённые на следующую строку (Clustered и прочие).
function Get-SectionEntries {
    param([string[]] $Lines, [string] $Section)
    $res = New-Object System.Collections.Generic.List[string]
    $head = '  ' + $Section
    $i = 0; $n = $Lines.Count
    while ($i -lt $n -and $Lines[$i].TrimEnd() -ne $head) { $i++ }
    if ($i -ge $n) { return ,$res }
    $i++
    if ($i -ge $n -or $Lines[$i].TrimEnd() -ne '  {') { return ,$res }
    $i++
    $cur = $null
    while ($i -lt $n) {
        $ln = $Lines[$i]
        if ($ln.TrimEnd() -eq '  }') { break }
        if ($ln -match '^    \{') {
            if ($null -ne $cur) { $res.Add($cur.ToString()) }
            $cur = New-Object System.Text.StringBuilder
            [void]$cur.Append($ln.Trim())
        }
        elseif ($null -ne $cur -and $ln.Trim().Length -gt 0) {
            [void]$cur.Append(' ')
            [void]$cur.Append($ln.Trim())
        }
        $i++
    }
    if ($null -ne $cur) { $res.Add($cur.ToString()) }
    return ,$res
}

# Тело записи без внешних фигурных скобок.
function Get-EntryBody {
    param([string] $Entry)
    $s = $Entry.Trim()
    if ($s.StartsWith('{')) { $s = $s.Substring(1) }
    $s = $s.TrimEnd()
    if ($s.EndsWith('}'))   { $s = $s.Substring(0, $s.Length - 1) }
    $s
}

# Значение свойства: от 'Имя=' до ';' или '}' на нулевой глубине скобок.
# Ретроспективный поиск обязателен: без него SQLIndex= ловится внутри MaintainSQLIndex=.
function Get-PropValue {
    param([string] $Text, [string] $Name)
    $m = [regex]::Match($Text, '(?<![A-Za-z])' + [regex]::Escape($Name) + '=')
    if (-not $m.Success) { return '' }
    $i = $m.Index + $m.Length
    $depth = 0
    $sb = New-Object System.Text.StringBuilder
    while ($i -lt $Text.Length) {
        $ch = $Text[$i]
        if ($ch -eq '(' -or $ch -eq '[') { $depth++ }
        elseif ($ch -eq ')' -or $ch -eq ']') { if ($depth -gt 0) { $depth-- } }
        elseif ($depth -eq 0 -and ($ch -eq ';' -or $ch -eq '}')) { break }
        [void]$sb.Append($ch)
        $i++
    }
    $sb.ToString().Trim()
}

# В TSV не должно попадать ничего, что ломает разбор по колонкам.
function Get-TsvSafe {
    param([string] $Value)
    if ($null -eq $Value) { return '' }
    ($Value -replace "[`t`r`n]", ' ').Trim()
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$ci = [System.Globalization.CultureInfo]::InvariantCulture

$parsed   = @{}   # TableID -> запись с самыми свежими Date/Time
$stale    = 0     # устаревшие файлы переименованных таблиц
$noHeader = 0     # файлы без строки OBJECT Table
$anomKey  = 0     # записи ключей, не разобранные на сегменты
$anomFld  = 0     # записи полей без числового номера
$done     = 0

foreach ($f in $files) {
    $lines = [System.IO.File]::ReadAllLines($f.FullName, [System.Text.Encoding]::UTF8)

    $tid = $null; $tname = ''
    $limit = [Math]::Min(5, $lines.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        if ($lines[$i] -match '^OBJECT Table (\d+) (.*?)\s*$') {
            $tid = [int]$Matches[1]; $tname = $Matches[2]; break
        }
    }
    if ($null -eq $tid) { $noHeader++; continue }

    # штамп объекта — по нему отсеиваются устаревшие файлы переименованных таблиц
    $stamp = [datetime]::MinValue
    $d = ''; $t = ''
    $limit = [Math]::Min(12, $lines.Count)
    for ($i = 0; $i -lt $limit; $i++) {
        if ($lines[$i] -match '^\s*Date=([\d.]+);') { $d = $Matches[1] }
        if ($lines[$i] -match '^\s*Time=([\d:]+);') { $t = $Matches[1] }
    }
    if ($d -and $t) {
        $tmp = [datetime]::MinValue
        if ([datetime]::TryParseExact("$d $t", 'dd.MM.yy HH:mm:ss', $ci, 'None', [ref]$tmp)) { $stamp = $tmp }
    }

    $prev = $parsed[$tid]
    if ($null -ne $prev) {
        $stale++
        if ($prev.Stamp -ge $stamp) { $done++; continue }
    }

    # --- ключи -------------------------------------------------------------
    $keyRows = New-Object System.Collections.Generic.List[string]
    $keyNo = 0; $keyOff = 0; $keySif = 0; $keyClust = 0
    foreach ($e in (Get-SectionEntries -Lines $lines -Section 'KEYS')) {
        $keyNo++
        $body = Get-EntryBody $e
        # сегменты: <Enabled> ; <KeyFields> ; <свойства>
        $seg = $body -split ';', 3
        if ($seg.Count -lt 2) { $anomKey++; continue }
        $enabled = if ($seg[0].Trim() -eq 'No') { 'No' } else { 'Yes' }
        $kf      = $seg[1].Trim()
        $props   = if ($seg.Count -ge 3) { $seg[2] } else { '' }

        $sif   = Get-PropValue $props 'SumIndexFields'
        $clust = Get-PropValue $props 'Clustered'
        $msql  = Get-PropValue $props 'MaintainSQLIndex'
        $msift = Get-PropValue $props 'MaintainSIFTIndex'
        $sqlix = Get-PropValue $props 'SQLIndex'

        if ($enabled -eq 'No')  { $keyOff++ }
        if ($sif)               { $keySif++ }
        if ($clust -eq 'Yes')   { $keyClust++ }

        $keyRows.Add((@(
            $tid
            (Get-TsvSafe $tname)
            $keyNo
            $enabled
            (Get-TsvSafe $kf)
            (Get-TsvSafe $sif)
            (Get-TsvSafe $clust)
            (Get-TsvSafe $msql)
            (Get-TsvSafe $msift)
            (Get-TsvSafe $sqlix)
        ) -join "`t"))
    }

    # --- поля --------------------------------------------------------------
    $fldRows = New-Object System.Collections.Generic.List[string]
    $fldOff = 0; $fldFlow = 0; $fldFilt = 0
    foreach ($e in (Get-SectionEntries -Lines $lines -Section 'FIELDS')) {
        $body = Get-EntryBody $e
        # сегменты: <No.> ; <Enabled> ; <Имя> ; <Тип> ; <свойства>
        $seg = $body -split ';', 5
        if ($seg.Count -lt 4 -or $seg[0].Trim() -notmatch '^\d+$') { $anomFld++; continue }
        $fno     = [int]$seg[0].Trim()
        $enabled = if ($seg[1].Trim() -eq 'No') { 'No' } else { 'Yes' }
        $fname   = $seg[2].Trim()
        $ftype   = $seg[3].Trim()
        $props   = if ($seg.Count -ge 5) { $seg[4] } else { '' }

        $fclass = Get-PropValue $props 'FieldClass'
        if (-not $fclass) { $fclass = 'Normal' }
        $formula = Get-PropValue $props 'CalcFormula'

        if ($enabled -eq 'No')         { $fldOff++ }
        if ($fclass -eq 'FlowField')   { $fldFlow++ }
        if ($fclass -eq 'FlowFilter')  { $fldFilt++ }

        $fldRows.Add((@(
            $tid
            (Get-TsvSafe $tname)
            $fno
            (Get-TsvSafe $fname)
            (Get-TsvSafe $ftype)
            (Get-TsvSafe $fclass)
            (Get-TsvSafe $formula)
            $enabled
        ) -join "`t"))
    }

    $parsed[$tid] = [pscustomobject]@{
        Id      = $tid
        Name    = $tname
        Stamp   = $stamp
        Keys    = $keyRows
        Fields  = $fldRows
        KeysOff = $keyOff
        KeysSif = $keySif
        KeysClu = $keyClust
        FldOff  = $fldOff
        FldFlow = $fldFlow
        FldFilt = $fldFilt
    }

    $done++
    if (-not $Quiet -and ($done % 500) -eq 0) {
        Write-Host ("  разобрано {0}..." -f $done) -ForegroundColor DarkGray
    }
}

# --- сборка файлов ----------------------------------------------------------
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Write-Tsv {
    param([string] $Path, [string] $Header, [System.Collections.Generic.List[string]] $Rows)
    $text = $Header + "`r`n"
    if ($Rows.Count -gt 0) { $text += (($Rows -join "`r`n") + "`r`n") }
    [System.IO.File]::WriteAllText($Path, $text, $utf8NoBom)
}

$ids = @($parsed.Keys | Sort-Object)

$keysAll = New-Object System.Collections.Generic.List[string]
$fldsAll = New-Object System.Collections.Generic.List[string]
$mapAll  = New-Object System.Collections.Generic.List[string]
$sqlMap  = @{}          # SqlName -> список 'id|name'
$totKeys = 0; $offKeys = 0; $sifKeys = 0; $cluKeys = 0
$totFlds = 0; $offFlds = 0; $flowFlds = 0; $filtFlds = 0
$noClustered = New-Object System.Collections.Generic.List[string]

foreach ($id in $ids) {
    $tbl = $parsed[$id]
    foreach ($r in $tbl.Keys)   { $keysAll.Add($r) }
    foreach ($r in $tbl.Fields) { $fldsAll.Add($r) }
    $totKeys += $tbl.Keys.Count; $offKeys  += $tbl.KeysOff; $sifKeys  += $tbl.KeysSif; $cluKeys += $tbl.KeysClu
    $totFlds += $tbl.Fields.Count; $offFlds += $tbl.FldOff;  $flowFlds += $tbl.FldFlow; $filtFlds += $tbl.FldFilt
    if ($tbl.KeysClu -eq 0 -and $tbl.Keys.Count -gt 0) {
        $noClustered.Add(("{0} {1}" -f $tbl.Id, $tbl.Name))
    }

    $sqlName = ConvertTo-SqlName $tbl.Name
    if (-not $sqlMap.ContainsKey($sqlName)) {
        $sqlMap[$sqlName] = New-Object System.Collections.Generic.List[string]
    }
    $sqlMap[$sqlName].Add(("{0}|{1}" -f $tbl.Id, $tbl.Name))
    $mapAll.Add((@((Get-TsvSafe $sqlName), $tbl.Id, (Get-TsvSafe $tbl.Name)) -join "`t"))
}

Write-Tsv (Join-Path $OutDir 'keys.tsv') `
    (@('TableID','TableName','KeyNo','Enabled','KeyFields','SumIndexFields','Clustered','MaintainSQLIndex','MaintainSIFTIndex','SQLIndex') -join "`t") `
    $keysAll
Write-Tsv (Join-Path $OutDir 'fields.tsv') `
    (@('TableID','TableName','FieldNo','FieldName','DataType','FieldClass','CalcFormula','Enabled') -join "`t") `
    $fldsAll
Write-Tsv (Join-Path $OutDir 'sqlmap.tsv') `
    (@('SqlName','TableID','TableName') -join "`t") `
    $mapAll

$collisions = @($sqlMap.Keys | Where-Object { $sqlMap[$_].Count -gt 1 })

# --- сверка словаря с фактическими именами в SQL ----------------------------
$sqlTotal = 0; $sqlMatched = 0; $sqlMissed = 0
$missedSample = New-Object System.Collections.Generic.List[string]
$sqlError = ''

if (-not $SkipSqlCheck) {
    $connStr = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True"
    $cn = New-Object System.Data.SqlClient.SqlConnection $connStr
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandTimeout = 120

        # префиксы компаний: в SQL имя таблицы имеет вид <Компания>$<Таблица>
        $prefixes = New-Object System.Collections.Generic.List[string]
        $cmd.CommandText = 'SELECT [Name] FROM dbo.[Company]'
        $rd = $cmd.ExecuteReader()
        while ($rd.Read()) { $prefixes.Add((ConvertTo-SqlName ([string]$rd['Name'])) + '$') }
        $rd.Close()
        $prefixes = @($prefixes | Sort-Object -Property Length -Descending)

        $cmd.CommandText = @"
SELECT t.[name] AS n
FROM   sys.tables t
JOIN   sys.schemas s ON s.schema_id = t.schema_id
WHERE  s.[name] = 'dbo' AND t.[name] NOT LIKE '`$ndo`$%'
ORDER BY t.[name]
"@
        $rd = $cmd.ExecuteReader()
        while ($rd.Read()) {
            $sqlTotal++
            $n = [string]$rd['n']
            foreach ($p in $prefixes) {
                if ($n.Length -gt $p.Length -and $n.StartsWith($p, [StringComparison]::OrdinalIgnoreCase)) {
                    $n = $n.Substring($p.Length); break
                }
            }
            if ($sqlMap.ContainsKey($n)) { $sqlMatched++ }
            else {
                $sqlMissed++
                if ($missedSample.Count -lt 8) { $missedSample.Add($n) }
            }
        }
        $rd.Close()
    }
    catch {
        $sqlError = $_.Exception.Message
    }
    finally {
        if ($cn.State -ne 'Closed') { $cn.Close() }
    }
}

$sw.Stop()

# --- итог (компактно, под один экран) ---------------------------------------
$ruleText = if ($SpaceToUnderscore) { '. " \ / '' % ] [ и пробел' } else { '. " \ / '' % ] [ , пробел сохраняется' }

Write-Host ''
Write-Host ("Экспорт:    {0}" -f $BaseDir)
Write-Host ("Каталог:    {0}" -f $OutDir)
Write-Host ("Файлов:     {0} (устаревших дублей {1}, без заголовка {2})" -f $files.Count, $stale, $noHeader)
Write-Host ("Таблиц:     {0}" -f $ids.Count)
Write-Host ("Ключей:     {0} (выключено {1}, SumIndexFields {2}, Clustered {3})" -f $totKeys, $offKeys, $sifKeys, $cluKeys)
Write-Host ("Полей:      {0} (выключено {1}, FlowField {2}, FlowFilter {3})" -f $totFlds, $offFlds, $flowFlds, $filtFlds)
Write-Host ("SQL-имена:  {0} записей, коллизий {1}; замена: {2}" -f $mapAll.Count, $collisions.Count, $ruleText)
if ($collisions.Count -gt 0) {
    $collisions | Select-Object -First 5 | ForEach-Object {
        Write-Host ("  {0} <- {1}" -f $_, (($sqlMap[$_]) -join ', ')) -ForegroundColor DarkYellow
    }
}
if ($SkipSqlCheck) {
    Write-Host 'Сверка:     пропущена (-SkipSqlCheck)' -ForegroundColor DarkGray
}
elseif ($sqlError) {
    Write-Host ("Сверка:     не выполнена — {0}" -f $sqlError) -ForegroundColor Yellow
}
else {
    $pct = 0.0
    if ($sqlTotal -gt 0) { $pct = 100.0 * $sqlMatched / $sqlTotal }
    Write-Host ("Сверка:     {0}\{1} — таблиц {2}, сопоставлено {3} ({4:N1} %), без пары {5}" -f `
        $Server, $Database, $sqlTotal, $sqlMatched, $pct, $sqlMissed)
    if ($missedSample.Count -gt 0) {
        Write-Host ("  без пары: {0}" -f ($missedSample -join ', ')) -ForegroundColor DarkYellow
    }
}
if ($anomKey -gt 0 -or $anomFld -gt 0) {
    Write-Host ("Аномалии:   ключей {0}, полей {1}" -f $anomKey, $anomFld) -ForegroundColor Yellow
}
if ($noClustered.Count -gt 0) {
    Write-Host ("Без Clustered: {0} табл. — {1}" -f $noClustered.Count, (($noClustered | Select-Object -First 3) -join '; ')) -ForegroundColor Yellow
}
Write-Host ("Время:      {0:N1} с" -f $sw.Elapsed.TotalSeconds)
