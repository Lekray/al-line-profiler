#Requires -Version 5.1
<#
.SYNOPSIS
    Прогон советника по ключам (Lib-KeyAdvisor) по трём корпусам запросов.

.DESCRIPTION
    В базу ничего не пишет: советник только читает справочники ключей и полей,
    собранные Build-KeysIndex.ps1 в <корень репозитория>\out.

    Корпус А — боевые тексты SQL из событий Application/705 «долгий SQL»
      (файл sql-corpus.tsv, готовит Export-SqlCorpus.ps1).
    Корпус Б — имена объектов SQL рабочей базы из sys.tables/sys.views
      (файл sysobjects.tsv); проверяется только сопоставление имени с таблицей NAV.
    Корпус В — запросы, собранные из справочников в форме, которую порождает NAV,
      с заранее известным ожидаемым вердиктом.

    Корпуса А и Б читают внешние файлы: их снимают с конкретной установки, и в
    репозитории их нет. Нет файла — корпус пропускается, остальные идут.

.PARAMETER DataDir
    Каталог корпусов и выходных файлов. По умолчанию <корень репозитория>\out.

.EXAMPLE
    powershell -File scripts\Test-KeyAdvisor.ps1
#>
[CmdletBinding()]
param(
    [string] $DataDir
)
$ErrorActionPreference = 'Stop'
$scripts = $PSScriptRoot
if (-not $DataDir) { $DataDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'out' }
. (Join-Path $scripts 'Lib-KeyAdvisor.ps1')
Initialize-KeyAdvisor

$Q = [char]34   # двойная кавычка

function Wrap-Sql {
    param([string]$Name)
    return ('SELECT * FROM {0}NAV{0}.dbo.{0}{1}{0} WITH(READUNCOMMITTED)' -f $Q, $Name)
}

function New-NavSql {
    <#  Шаблон снят с формы реального прикладного запроса из журнала:
        SELECT TOP (@0) "18"."No_" FROM "NAV".dbo."NAV$Customer" "18"
        WITH(READUNCOMMITTED) WHERE ("18"."No_"<@1) ORDER BY "No_" DESC
        OPTION(OPTIMIZE FOR UNKNOWN, FAST 50) AppObjectType: Page AppObjectId: 22  #>
    param(
        [int]$TableId, [string]$SqlTable, [string[]]$Where, [string[]]$OrderBy,
        [string]$Select, [string]$Company = 'NAV', [switch]$NoPrefix, [string]$ViewSuffix
    )
    $a = [string]$TableId
    $obj = $SqlTable
    if (-not $NoPrefix) { $obj = $Company + '$' + $SqlTable }
    if ($ViewSuffix)    { $obj = $obj + $ViewSuffix }
    $sel = $Select
    if (-not $sel) { $sel = 'TOP (@0) {0}{1}{0}.*' -f $Q, $a }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(('SELECT {0} FROM {1}NAV{1}.dbo.{1}{2}{1} {1}{3}{1} WITH(READUNCOMMITTED)' -f $sel, $Q, $obj, $a))
    if (@($Where).Count -gt 0)   { [void]$sb.Append(' WHERE ' + (@($Where) -join ' AND ')) }
    if (@($OrderBy).Count -gt 0) { [void]$sb.Append(' ORDER BY ' + ((@($OrderBy) | ForEach-Object { '{0}{1}{0}' -f $Q, $_ }) -join ',')) }
    [void]$sb.Append(' OPTION(OPTIMIZE FOR UNKNOWN, FAST 50) AppObjectType: Codeunit AppObjectId: 80')
    return $sb.ToString()
}

function Col { param([int]$TableId, [string]$Field) return ('{0}{1}{0}.{0}{2}{0}' -f $Q, $TableId, (ConvertTo-KaSqlName $Field)) }

# ===========================================================================
# Корпус А — боевой: события Application/705 «Long running SQL statement»
# ===========================================================================
Write-Host ''
Write-Host '=== КОРПУС А: журнал Windows, события 705 ===' -ForegroundColor Cyan
$corpusA = Join-Path $DataDir 'sql-corpus.tsv'
if (-not (Test-Path -LiteralPath $corpusA)) {
    Write-Host ('Пропущен: нет файла ' + $corpusA) -ForegroundColor DarkYellow
    Write-Host '  Это выгрузка боевых текстов SQL из журнала Windows (Application, событие 705).' -ForegroundColor DarkYellow
    Write-Host '  Собрать: powershell -File scripts\Export-SqlCorpus.ps1 -OutFile <этот путь>' -ForegroundColor DarkYellow
}
else {
    $rows = [System.IO.File]::ReadAllLines($corpusA, [Text.Encoding]::UTF8)
    $sqls = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -lt $rows.Length; $i++) {
        if (-not $rows[$i]) { continue }
        $c = $rows[$i] -split "`t"
        if ($c.Length -ge 5) { $sqls.Add($c[4]) }
    }
    $uniq = @($sqls | Select-Object -Unique)
    Write-Host ("Запросов: {0}, различных текстов: {1}" -f $sqls.Count, $uniq.Count)

    $statA = @{}; $tblOk = 0; $ambig = 0; $silent = 0; $unres = 0; $noWhere = 0; $unknownPred = 0
    $byText = @{}
    foreach ($s in $uniq) {
        $v = Invoke-KeyAdvisor -Sql $s
        $byText[$s] = $v
        if (-not $statA.ContainsKey($v.Verdict)) { $statA[$v.Verdict] = 0 }
        $statA[$v.Verdict]++
    }
    foreach ($s in $sqls) {
        $v = $byText[$s]
        if ($v.Table.Ok)                 { $tblOk++ }
        if ($v.Verdict -eq 'Ambiguous')  { $ambig++ }
        if ($v.Verdict -eq 'Silent')     { $silent++ }
        if ($v.Verdict -eq 'Unresolved') { $unres++ }
        if ($v.Predicates -and -not $v.Predicates.HasWhere) { $noWhere++ }
        if ($v.Predicates) { $unknownPred += @($v.Predicates.Unknown).Count }
    }
    Write-Host ("Таблица определена:  {0} из {1} запросов" -f $tblOk, $sqls.Count)
    Write-Host ("Служебные (молчим):  {0}; неразобранных: {1}; неоднозначных: {2}" -f $silent, $unres, $ambig)
    Write-Host ("Без WHERE:           {0}; предикатов Unknown: {1}" -f $noWhere, $unknownPred)
    Write-Host 'Вердикты по различным текстам:'
    foreach ($k in ($statA.Keys | Sort-Object)) { Write-Host ("  {0,-16} {1}" -f $k, $statA[$k]) }
}

# ===========================================================================
# Корпус Б — реальные имена объектов SQL из sys.tables/sys.views рабочей базы
# ===========================================================================
Write-Host ''
Write-Host '=== КОРПУС Б: реальные имена объектов SQL рабочей базы ===' -ForegroundColor Cyan
$corpusB = Join-Path $DataDir 'sysobjects.tsv'
if (-not (Test-Path -LiteralPath $corpusB)) {
    Write-Host ('Пропущен: нет файла ' + $corpusB) -ForegroundColor DarkYellow
    Write-Host '  Это выгрузка имён объектов SQL целевой базы: sys.tables и sys.views,' -ForegroundColor DarkYellow
    Write-Host '  по одному имени в строке, имя — первая колонка через табуляцию.' -ForegroundColor DarkYellow
    Write-Host '  Снять: SELECT name FROM sys.tables UNION ALL SELECT name FROM sys.views' -ForegroundColor DarkYellow
    Write-Host '  В репозиторий такой файл не кладётся: это имена конкретной установки.' -ForegroundColor DarkYellow
}
else {
    $names = [System.IO.File]::ReadAllLines($corpusB, [Text.Encoding]::UTF8)
    $res = @{ Resolved = 0; System = 0; NotFound = 0; Ambiguous = 0; Sift = 0 }
    $missed = New-Object System.Collections.Generic.List[string]
    foreach ($ln in $names) {
        if (-not $ln) { continue }
        $n = ($ln -split "`t")[0]
        $r = Get-SqlTableRef -Sql (Wrap-Sql $n)
        if ($r.IsSift) { $res.Sift++ }
        switch ($r.Status) {
            'Resolved'  { $res.Resolved++ }
            'System'    { $res.System++ }
            'Ambiguous' { $res.Ambiguous++ }
            default     { $res.NotFound++; if ($missed.Count -lt 12) { $missed.Add($n) } }
        }
    }
    $tot = $res.Resolved + $res.System + $res.NotFound + $res.Ambiguous
    Write-Host ("Всего имён: {0}" -f $tot)
    Write-Host ("  сопоставлено с таблицей NAV: {0} ({1:N1} %)" -f $res.Resolved, (100.0 * $res.Resolved / $tot))
    Write-Host ("  служебные платформы:         {0}" -f $res.System)
    Write-Host ("  неоднозначных:               {0}" -f $res.Ambiguous)
    Write-Host ("  не найдено в словаре:        {0}" -f $res.NotFound)
    Write-Host ("  из них представлений SIFT распознано: {0}" -f $res.Sift)
    if ($missed.Count -gt 0) { Write-Host ('  примеры ненайденных: ' + ($missed -join '; ')) -ForegroundColor DarkYellow }
}

# ===========================================================================
# Корпус В — запросы, собранные по МЕТАДАННЫМ рабочей базы в форме реального запроса NAV
# ===========================================================================
Write-Host ''
Write-Host '=== КОРПУС В: запросы по метаданным рабочей базы в форме NAV ===' -ForegroundColor Cyan

$cases = New-Object System.Collections.ArrayList
function Add-Case {
    param([string]$Name, [string]$Sql, [string]$Expect)
    [void]$cases.Add([pscustomobject]@{ Name = $Name; Sql = $Sql; Expect = $Expect })
}

# --- поиск подходящих таблиц прямо в справочнике ---------------------------
$allKeys = $script:KaKeys
function Get-KeyOf { param([int]$T,[int]$N) return @(Get-KaTableKeys $T | Where-Object { $_.KeyNo -eq $N })[0] }

# 1. ключ ложится целиком: берём включённый ключ из >=3 полей
$src = $null
foreach ($tid in @(32, 17, 21, 5802)) {
    $k = @(Get-KaTableKeys $tid | Where-Object { $_.Enabled -and $_.MaintainSQLIndex -and @($_.Fields).Count -ge 3 -and @($_.SqlIndexFields).Count -eq 0 })
    if ($k.Count -gt 0) { $src = $k[0]; break }
}
if ($src) {
    $f = @($src.Fields)
    $e = @($f[0..($f.Count - 3)])
    $r = $f[$f.Count - 2]
    $w = @(); foreach ($x in $e) { $w += ('{0}=@{1}' -f (Col $src.TableId $x), ($w.Count + 1)) }
    $w += ('{0}>=@9' -f (Col $src.TableId $r))
    Add-Case ("ключ {0} таблицы {1} ложится целиком" -f $src.KeyNo, $src.TableId) `
        (New-NavSql -TableId $src.TableId -SqlTable (ConvertTo-KaSqlName $src.TableName) -Where $w -OrderBy (@($f) | ForEach-Object { ConvertTo-KaSqlName $_ })) 'KeyExists'
}

# 2. выключенный ключ, состав которого не покрыт ни одним включённым
$dis = $null
foreach ($tid in ($script:KaKeys.Keys | Sort-Object)) {
    if ($tid -lt 1 -or $tid -ge 2000000000) { continue }
    $ks = @(Get-KaTableKeys $tid)
    $off = @($ks | Where-Object { -not $_.Enabled -and @($_.Fields).Count -ge 2 })
    foreach ($k in $off) {
        $eL = @($k.FieldsLower)
        $covered = $false
        foreach ($o in $ks) {
            if (-not $o.Enabled) { continue }
            $of = @($o.EffectiveLower)
            $m = 0; while ($m -lt $of.Count -and ($eL -contains $of[$m])) { $m++ }
            if ($m -eq $eL.Count) { $covered = $true; break }
        }
        if (-not $covered) { $dis = $k; break }
    }
    if ($dis) { break }
}
if ($dis) {
    $w = @(); foreach ($x in @($dis.Fields)) { $w += ('{0}=@{1}' -f (Col $dis.TableId $x), ($w.Count + 1)) }
    Add-Case ("выключенный ключ {0} таблицы {1}" -f $dis.KeyNo, $dis.TableId) `
        (New-NavSql -TableId $dis.TableId -SqlTable (ConvertTo-KaSqlName $dis.TableName) -Where $w) 'KeyDisabled'
}

# 3. ключ с MaintainSQLIndex=No
$nom = $null
foreach ($tid in ($script:KaKeys.Keys | Sort-Object)) {
    if ($tid -lt 1 -or $tid -ge 2000000000) { continue }
    foreach ($k in (Get-KaTableKeys $tid)) {
        if ($k.Enabled -and -not $k.MaintainSQLIndex -and @($k.Fields).Count -ge 2) { $nom = $k; break }
    }
    if ($nom) { break }
}
if ($nom) {
    $w = @(); foreach ($x in @($nom.Fields)) { $w += ('{0}=@{1}' -f (Col $nom.TableId $x), ($w.Count + 1)) }
    Add-Case ("ключ {0} таблицы {1} с MaintainSQLIndex=No" -f $nom.KeyNo, $nom.TableId) `
        (New-NavSql -TableId $nom.TableId -SqlTable (ConvertTo-KaSqlName $nom.TableName) -Where $w) 'KeyNotMaintained'
}

# 4. удлинить: ключ целиком лежит внутри фильтра, не хватает одного поля
$ext = $null
foreach ($tid in ($script:KaKeys.Keys | Sort-Object)) {
    if ($tid -lt 1 -or $tid -ge 2000000000) { continue }
    $ks = @(Get-KaTableKeys $tid)
    $short = @($ks | Where-Object { $_.Enabled -and $_.MaintainSQLIndex -and @($_.Fields).Count -eq 1 -and @($_.SqlIndexFields).Count -eq 0 })
    if ($short.Count -eq 0) { continue }
    $k = $short[0]
    $extra = @($script:KaFields[$tid].Values | Where-Object { $_.FieldClass -eq 'Normal' -and $_.Enabled } |
              Sort-Object FieldNo -Unique | Where-Object { @($k.FieldsLower) -notcontains $_.SqlName.ToLowerInvariant() })
    foreach ($x in $extra) {
        $eL = @(@($k.FieldsLower) + @($x.SqlName.ToLowerInvariant()))
        $covered = $false
        foreach ($o in $ks) {
            $of = @($o.EffectiveLower); $m = 0
            while ($m -lt $of.Count -and ($eL -contains $of[$m])) { $m++ }
            if ($m -eq $eL.Count) { $covered = $true; break }
        }
        if (-not $covered) { $ext = [pscustomobject]@{ Key = $k; Extra = $x.Name }; break }
    }
    if ($ext) { break }
}
if ($ext) {
    $tid = $ext.Key.TableId
    $w = @(('{0}=@1' -f (Col $tid (@($ext.Key.Fields)[0]))), ('{0}=@2' -f (Col $tid $ext.Extra)))
    Add-Case ('удлинить ключ {0} таблицы {1}: не хватает одного поля' -f $ext.Key.KeyNo, $tid) `
        (New-NavSql -TableId $tid -SqlTable (ConvertTo-KaSqlName $ext.Key.TableName) -Where $w) 'ExtendKey'
}

# 5. новый ключ: клиентская таблица 50000+, фильтр по полям, с которых не начинается ни один ключ
$cust = $null
foreach ($tid in ($script:KaKeys.Keys | Sort-Object)) {
    if ($tid -lt 50000 -or $tid -gt 99999) { continue }
    $ks = @(Get-KaTableKeys $tid)
    $flds = @($script:KaFields[$tid].Values | Where-Object { $_.FieldClass -eq 'Normal' -and $_.Enabled } | Sort-Object FieldNo -Unique)
    $firsts = @($ks | ForEach-Object { @($_.EffectiveLower)[0] })
    $cand = @($flds | Where-Object { $firsts -notcontains $_.SqlName.ToLowerInvariant() })
    if ($cand.Count -ge 2 -and $ks.Count -ge 1) {
        $cust = [pscustomobject]@{ TableId = $tid; Name = $ks[0].TableName; F1 = $cand[0].Name; F2 = $cand[1].Name }
        break
    }
}
if ($cust) {
    $w = @(('{0}=@1' -f (Col $cust.TableId $cust.F1)), ('{0}=@2' -f (Col $cust.TableId $cust.F2)))
    Add-Case ("новый ключ на клиентской таблице {0}" -f $cust.TableId) `
        (New-NavSql -TableId $cust.TableId -SqlTable (ConvertTo-KaSqlName $cust.Name) -Where $w) 'NewKey'
}

# 6. агрегат мимо SIFT: SUM по полю, которого нет ни в одном SumIndexFields
$noSift = $null
foreach ($tid in @(32, 17, 5802, 21)) {
    $sums = @()
    foreach ($k in (Get-KaTableKeys $tid)) { $sums += @($k.SumLower) }
    $dec = @($script:KaFields[$tid].Values | Where-Object { $_.DataType -like 'Decimal*' -and $_.FieldClass -eq 'Normal' } |
             Where-Object { $sums -notcontains $_.SqlName.ToLowerInvariant() } | Sort-Object FieldNo -Unique)
    if ($dec.Count -gt 0) {
        $k1 = Get-KeyOf $tid 2
        if ($k1) { $noSift = [pscustomobject]@{ TableId = $tid; Name = $k1.TableName; Field = $dec[0].Name; KeyField = @($k1.Fields)[0] }; break }
    }
}
if ($noSift) {
    $sel = 'SUM({0}{1}{0}.{0}{2}{0})' -f $Q, $noSift.TableId, (ConvertTo-KaSqlName $noSift.Field)
    $w = @(('{0}=@1' -f (Col $noSift.TableId $noSift.KeyField)))
    Add-Case ("агрегат мимо SIFT по таблице {0}" -f $noSift.TableId) `
        (New-NavSql -TableId $noSift.TableId -SqlTable (ConvertTo-KaSqlName $noSift.Name) -Where $w -Select $sel) 'KeyExists'
}

# 7. агрегат через представление SIFT
$k2s = Get-KeyOf 32 2
if ($k2s) {
    $sel = 'SUM({0}32{0}.{0}SUM$Quantity{0})' -f $Q
    $w = @(('{0}=@1' -f (Col 32 (@($k2s.Fields)[0]))))
    Add-Case 'агрегат через представление SIFT (VSIFT$1)' `
        (New-NavSql -TableId 32 -SqlTable 'Item Ledger Entry' -Where $w -Select $sel -ViewSuffix '$VSIFT$1') 'KeyExists'
}

# 8. SIFT в обход: SUM по полю, которое ЕСТЬ в SumIndexFields, но читается базовая таблица
if ($k2s) {
    $sums = @($k2s.SumIndexFields)
    if ($sums.Count -gt 0) {
        $sel = 'SUM({0}32{0}.{0}{1}{0})' -f $Q, (ConvertTo-KaSqlName $sums[0])
        $w = @(('{0}=@1' -f (Col 32 'Document No.')))
        Add-Case 'SIFT в обход: сумма по базовой таблице' `
            (New-NavSql -TableId 32 -SqlTable 'Item Ledger Entry' -Where $w -Select $sel) 'любой'
    }
}

# 9-16. формы предикатов: что попадает в N и в Unknown
$T2 = 32
$c  = { param($f) Col $T2 $f }
Add-Case 'неравенство <>' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0}=@1' -f (& $c 'Item No.')), ('{0}<>@2' -f (& $c 'Location Code')))) 'KeyExists'
Add-Case 'LIKE с маской слева' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ("{0} LIKE 'A%B'" -f (& $c 'Item No.') -replace 'A%B', "%$([char]37)ABC$([char]37)"))) 'NoIndexablePredicate'
Add-Case 'LIKE с параметром (форма неизвестна)' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0} LIKE @1' -f (& $c 'Item No.')))) 'NoIndexablePredicate'
Add-Case 'функция над колонкой' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('UPPER({0})=@1' -f (& $c 'Item No.')))) 'NoIndexablePredicate'
Add-Case 'OR по разным полям' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('({0}=@1 OR {1}=@2)' -f (& $c 'Item No.'), (& $c 'Document No.')))) 'NoIndexablePredicate'
Add-Case 'OR по одному полю = набор значений' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('({0}=@1 OR {0}=@2)' -f (& $c 'Item No.')))) 'KeyExists'
Add-Case 'IN как равенство' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0} IN (@1,@2,@3)' -f (& $c 'Item No.')))) 'KeyExists'
Add-Case 'BETWEEN как диапазон' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0}=@1' -f (& $c 'Item No.')), ('{0} BETWEEN @2 AND @3' -f (& $c 'Posting Date')))) 'KeyExists'
Add-Case 'IS NULL как равенство' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0} IS NULL' -f (& $c 'Item No.')))) 'KeyExists'
Add-Case 'NOT IN не индексируется' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0}=@1' -f (& $c 'Item No.')), ('{0} NOT IN (@2,@3)' -f (& $c 'Location Code')))) 'KeyExists'

# 17-20. политика по классам таблиц
$core = @($script:KaTables.Keys |
          Where-Object { $_ -ge $script:KaRangeClientCoreFrom -and $_ -lt $script:KaRangeVirtualFrom } | Sort-Object)
if ($core.Count -gt 0) {
    Add-Case ('ядро отраслевого решения 10 000 000+ (таблица {0})' -f $core[0]) `
        (New-NavSql -TableId $core[0] -SqlTable (ConvertTo-KaSqlName $script:KaTables[$core[0]]) -Where @(
            ('{0}{1}{0}.{0}Some Field{0}=@1' -f $Q, $core[0]))) 'любой'
}
$virt = @($script:KaTables.Keys | Where-Object { $_ -ge 2000000000 } | Sort-Object)
if ($virt.Count -gt 0) {
    Add-Case ("виртуальная таблица {0}" -f $virt[0]) (New-NavSql -TableId $virt[0] -SqlTable (ConvertTo-KaSqlName $script:KaTables[$virt[0]]) -Where @('1=@1')) 'Silent'
}
Add-Case 'служебная таблица платформы' 'SELECT [id] FROM [NAV].[dbo].[$ndo$taskscheduling] WITH (READCOMMITTED) WHERE [timestamp] > @timestamp' 'Silent'
Add-Case 'имени нет в словаре, но псевдоним = номер таблицы' (New-NavSql -TableId 32 -SqlTable 'Not A Real Table Name' -Where @(
    ('{0}32{0}.{0}Item No_{0}=@1' -f $Q))) 'KeyExists'
Add-Case 'имя и псевдоним расходятся' (New-NavSql -TableId 17 -SqlTable 'Item Ledger Entry' -Where @(
    ('{0}17{0}.{0}Item No_{0}=@1' -f $Q))) 'Ambiguous'
Add-Case 'ни фильтра, ни сортировки' (New-NavSql -TableId $T2 -SqlTable 'Item Ledger Entry') 'NoPredicates'

# --- прогон -----------------------------------------------------------------
$stat = @{}; $okAssert = 0; $badAssert = 0
$bad = New-Object System.Collections.ArrayList
foreach ($cs in $cases) {
    $v = Invoke-KeyAdvisor -Sql $cs.Sql
    if (-not $stat.ContainsKey($v.Verdict)) { $stat[$v.Verdict] = 0 }
    $stat[$v.Verdict]++
    if ($cs.Expect -eq 'любой' -or $v.Verdict -eq $cs.Expect) { $okAssert++ }
    else { $badAssert++; [void]$bad.Add(('{0}: ждали {1}, получили {2}' -f $cs.Name, $cs.Expect, $v.Verdict)) }
}
Write-Host ("Случаев: {0}; ожидание совпало: {1}; разошлось: {2}" -f $cases.Count, $okAssert, $badAssert)
foreach ($k in ($stat.Keys | Sort-Object)) { Write-Host ("  {0,-16} {1}" -f $k, $stat[$k]) }
foreach ($b in $bad) { Write-Host ('  РАСХОЖДЕНИЕ: ' + $b) -ForegroundColor Yellow }

$sq = @($cases | ForEach-Object { $_.Sql })
$agg = 0; $noSiftN = 0; $bypass = 0; $viaSift = 0
foreach ($s in $sq) {
    $v = Invoke-KeyAdvisor -Sql $s
    if ($v.Sift -and $v.Sift.IsAggregate) {
        $agg++
        if ($v.Sift.Verdict -eq 'NoSift')      { $noSiftN++ }
        if ($v.Sift.Verdict -eq 'SiftBypassed'){ $bypass++ }
        if ($v.Sift.Verdict -eq 'ViaSift')     { $viaSift++ }
    }
}
Write-Host ("Агрегатов: {0}; без SIFT: {1}; SIFT в обход: {2}; через SIFT: {3}" -f $agg, $noSiftN, $bypass, $viaSift)

$dump = New-Object System.Collections.ArrayList
foreach ($cs in $cases) {
    $v = Invoke-KeyAdvisor -Sql $cs.Sql
    [void]$dump.Add('### ' + $cs.Name + '  [вердикт: ' + $v.Verdict + ']')
    [void]$dump.Add('SQL: ' + $cs.Sql)
    [void]$dump.Add($v.Advice)
    [void]$dump.Add('')
}
[System.IO.File]::WriteAllText((Join-Path $DataDir 'advice-samples.txt'), (($dump -join "`r`n") + "`r`n"), (New-Object Text.UTF8Encoding($false)))
Write-Host ('Примеры советов: ' + (Join-Path $DataDir 'advice-samples.txt'))
