#Requires -Version 5.1
<#
.SYNOPSIS
    Советник по ключам: по тексту SQL, который сгенерировал NAV, — вердикт о покрытии
    индексом и готовая формулировка совета. Подключается через dot-sourcing.

.DESCRIPTION
    Библиотека отвечает на один вопрос: «этот запрос упирается в отсутствие ключа
    или нет, и что с этим делать». Работает по трём справочникам из
    <корень репозитория>\out (их строит Build-KeysIndex.ps1):
        keys.tsv    TableID, TableName, KeyNo, Enabled, KeyFields, SumIndexFields,
                    Clustered, MaintainSQLIndex, MaintainSIFTIndex, SQLIndex
        fields.tsv  TableID, TableName, FieldNo, FieldName, DataType, FieldClass,
                    CalcFormula, Enabled
        sqlmap.tsv  SqlName, TableID, TableName

    Что здесь считается фактом, а не догадкой (проверено на рабочей базе):
      - имя SQL-индекса выводится из СКВОЗНОГО номера ключа, вместе с выключенными:
        ключ 1 (кластерный) -> "<SQL-имя таблицы>$0", ключ N -> "$(N-1)".
        На dbo.[NAV$G_L Entry] выключенные ключи 4 и 5 съедают имена $3 и $4,
        которых в sys.indexes нет;
      - индексированное представление SIFT называется "<префикс>$<таблица>$VSIFT$(N-1)"
        и существует ровно для ключей, у которых Enabled=Yes, SumIndexFields непусты
        и MaintainSIFTIndex<>No. Сверено на NAV$Item Ledger Entry: представления
        $VSIFT$1,3,4,5,6,15,22 против ключей 2,4,5,6,7,16,23 — совпадение полное;
      - имя SQL-объекта получается из имени NAV заменой символов . " \ / ' % ] [ ;
        ПРОБЕЛ не заменяется. То же преобразование применяется и к именам КОЛОНОК:
        поле «G/L Account No.» лежит в колонке [G_L Account No_];
      - префикс компании отделяется по первому $; таблицы с DataPerCompany=No
        лежат без префикса, служебные — с именем вида $ndo$…;
      - NAV даёт таблице псевдоним, равный НОМЕРУ таблицы: FROM "NAV".dbo."NAV$Customer"
        "18". Это независимая проверка разбора имени.

    Осторожность важнее полноты: если имя таблицы неоднозначно или предикат не разобран,
    библиотека сообщает об этом, а не угадывает. Поля из неиндексируемого множества
    в предлагаемый ключ не попадают никогда.

    В базу ничего не пишется, внешних зависимостей нет.

    Грабли Windows PowerShell 5.1 (проверено на 5.1.26100.9168): внутри функции
    выражение @(<System.Collections.Generic.List[object]>) падает с ArgumentException
    "Argument types do not match" — ломается связыватель PSToObjectArrayBinder.
    List[string] и ArrayList тем же местом не задеты, поэтому списки объектов здесь
    держатся в ArrayList, а Add глушится [void]. Если менять — не возвращать
    List[object] наружу.

    Подключение:  . (Join-Path $PSScriptRoot 'Lib-KeyAdvisor.ps1')

.EXAMPLE
    . .\Lib-KeyAdvisor.ps1
    $v = Invoke-KeyAdvisor -Sql $sql
    $v.Verdict; $v.Advice
#>

# ============================================================================
#  Справочники
# ============================================================================

$script:KaLoaded   = $false
$script:KaOutDir   = $null
$script:KaKeys     = $null   # [int]    TableID -> List[ключ]
$script:KaFields   = $null   # [int]    TableID -> Hashtable имя(lower) -> поле
$script:KaTables   = $null   # [int]    TableID -> имя таблицы
$script:KaSqlMap   = $null   # [string] SQL-имя(lower) -> List[кандидат]
$script:KaCache    = $null   # [string] текст SQL -> результат Invoke-KeyAdvisor

# Символы, которые платформа заменяет на подчёркивание в имени SQL-объекта и колонки.
$script:KaBadChars = '."\/' + "'" + '%]['

# Классы таблиц: границы диапазонов номеров.
$script:KaRangeCustomFrom     = 50000
$script:KaRangeCustomTo       = 99999
$script:KaRangeClientCoreFrom = 10000000
$script:KaRangeVirtualFrom    = 2000000000

function ConvertTo-KaSqlName {
    <#
    .SYNOPSIS
        Имя NAV -> имя SQL-объекта или колонки (пробел не трогаем).
    #>
    param([string] $Name)
    if (-not $Name) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($script:KaBadChars.IndexOf($ch) -ge 0) { [void]$sb.Append('_') }
        else                                       { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Split-KaFieldList {
    <#
    .SYNOPSIS
        Состав ключа в список полей: запятая делит, но НЕ внутри кавычек.

    .DESCRIPTION
        Имя поля само может содержать запятую, и тогда C/SIDE берёт его в кавычки:
        SumIndexFields=Quantity,"Qty. (Absolute, Base)". Наивное деление рвало такое
        имя пополам, обе половины уходили в справочник мусором, а дальше это
        оборачивалось и лишним полем в составе ключа, и «поле не объявлено в
        SumIndexFields» для поля, которое там объявлено.
    #>
    param([string] $Text)
    if (-not $Text) { return @() }
    $out  = New-Object System.Collections.Generic.List[string]
    $sb   = New-Object System.Text.StringBuilder
    $inQ  = $false
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($ch -eq '"') {
            # Удвоенная кавычка внутри значения - это сама кавычка, а не граница.
            if ($inQ -and ($i + 1) -lt $Text.Length -and $Text[$i + 1] -eq '"') {
                [void]$sb.Append('"'); $i++; continue
            }
            $inQ = -not $inQ
            continue
        }
        if ($ch -eq ',' -and -not $inQ) { [void]$out.Add($sb.ToString()); [void]$sb.Clear(); continue }
        [void]$sb.Append($ch)
    }
    [void]$out.Add($sb.ToString())
    return @($out | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-KaDefaultOutDir {
    # scripts -> корень репозитория
    return (Join-Path (Split-Path -Parent $PSScriptRoot) 'out')
}

function Read-KaTsv {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Не найден справочник $Path. Сначала выполните Build-KeysIndex.ps1."
    }
    return [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
}

function Initialize-KeyAdvisor {
    <#
    .SYNOPSIS
        Загружает keys.tsv, fields.tsv и sqlmap.tsv в память. Идемпотентна.

    .PARAMETER OutDir
        Каталог справочников. По умолчанию <корень репозитория>\out.

    .PARAMETER Force
        Перечитать, даже если справочники уже загружены.
    #>
    [CmdletBinding()]
    param(
        [string] $OutDir,
        [switch] $Force
    )
    if ($script:KaLoaded -and -not $Force) { return }
    if (-not $OutDir) { $OutDir = Get-KaDefaultOutDir }

    $keys   = @{}
    $fields = @{}
    $tables = @{}
    $sqlmap = @{}

    # --- поля ---------------------------------------------------------------
    $lines = Read-KaTsv (Join-Path $OutDir 'fields.tsv')
    for ($i = 1; $i -lt $lines.Length; $i++) {
        if (-not $lines[$i]) { continue }
        $c = $lines[$i] -split "`t"
        if ($c.Length -lt 8) { continue }
        $tid = [int]$c[0]
        if (-not $fields.ContainsKey($tid)) { $fields[$tid] = @{} }
        $fld = [pscustomobject]@{
            TableId     = $tid
            FieldNo     = [int]$c[2]
            Name        = $c[3]
            SqlName     = (ConvertTo-KaSqlName $c[3])
            DataType    = $c[4]
            FieldClass  = $c[5]
            CalcFormula = $c[6]
            Enabled     = ($c[7] -eq 'Yes')
        }
        # ключ и по имени NAV, и по имени колонки SQL — в запросе встречается второе
        $fields[$tid][$fld.Name.ToLowerInvariant()]    = $fld
        $fields[$tid][$fld.SqlName.ToLowerInvariant()] = $fld
        if (-not $tables.ContainsKey($tid)) { $tables[$tid] = $c[1] }
    }

    # --- ключи --------------------------------------------------------------
    $lines = Read-KaTsv (Join-Path $OutDir 'keys.tsv')
    for ($i = 1; $i -lt $lines.Length; $i++) {
        if (-not $lines[$i]) { continue }
        $c = $lines[$i] -split "`t"
        if ($c.Length -lt 10) { continue }
        $tid = [int]$c[0]
        if (-not $keys.ContainsKey($tid)) { $keys[$tid] = New-Object System.Collections.ArrayList }
        if (-not $tables.ContainsKey($tid)) { $tables[$tid] = $c[1] }

        $kf = @(Split-KaFieldList $c[4])
        $si = @(Split-KaFieldList $c[5])
        $sq = @(Split-KaFieldList $c[9])

        # пустое значение свойства означает умолчание Yes
        $maintainSql  = ($c[7] -ne 'No')
        $maintainSift = ($c[8] -ne 'No')
        $keyNo        = [int]$c[2]
        $enabled      = ($c[3] -eq 'Yes')

        # состав, по которому индекс реально живёт в SQL: SQLIndex перебивает состав ключа
        $eff = $kf
        if ($sq.Count -gt 0) { $eff = $sq }

        [void]$keys[$tid].Add([pscustomobject]@{
            TableId           = $tid
            TableName         = $c[1]
            KeyNo             = $keyNo
            Enabled           = $enabled
            Fields            = $kf
            SumIndexFields    = $si
            Clustered         = ($c[6] -eq 'Yes')
            MaintainSQLIndex  = $maintainSql
            MaintainSIFTIndex = $maintainSift
            SqlIndexFields    = $sq
            EffectiveFields   = $eff
            EffectiveLower    = @($eff | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() })
            FieldsLower       = @($kf  | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() })
            SumLower          = @($si  | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() })
            IndexSuffix       = ('$' + ($keyNo - 1))
            HasSift           = (($si.Count -gt 0) -and $enabled -and $maintainSift)
        })
    }

    # --- словарь имён -------------------------------------------------------
    $lines = Read-KaTsv (Join-Path $OutDir 'sqlmap.tsv')
    for ($i = 1; $i -lt $lines.Length; $i++) {
        if (-not $lines[$i]) { continue }
        $c = $lines[$i] -split "`t"
        if ($c.Length -lt 3) { continue }
        $k = $c[0].ToLowerInvariant()
        if (-not $sqlmap.ContainsKey($k)) { $sqlmap[$k] = New-Object System.Collections.ArrayList }
        [void]$sqlmap[$k].Add([pscustomobject]@{ TableId = [int]$c[1]; TableName = $c[2] })
        if (-not $tables.ContainsKey([int]$c[1])) { $tables[[int]$c[1]] = $c[2] }
    }

    $script:KaOutDir = $OutDir
    $script:KaKeys   = $keys
    $script:KaFields = $fields
    $script:KaTables = $tables
    $script:KaSqlMap = $sqlmap
    $script:KaCache  = @{}
    $script:KaLoaded = $true
}

function ConvertTo-KaFieldName {
    <#
    .SYNOPSIS
        Имя колонки SQL -> имя поля C/AL. В совет должно попадать имя, которое
        пишут в ключ: «Document No.», а не «Document No_».
    #>
    param([int] $TableId, [string] $Name)
    $f = Get-KaField -TableId $TableId -Name $Name
    if ($f) { return $f.Name }
    return $Name
}

function Get-KaTableKeys {
    param([int] $TableId)
    Initialize-KeyAdvisor
    if ($script:KaKeys.ContainsKey($TableId)) { return $script:KaKeys[$TableId] }
    return @()
}

function Get-KaField {
    param([int] $TableId, [string] $Name)
    Initialize-KeyAdvisor
    if (-not $script:KaFields.ContainsKey($TableId)) { return $null }
    $k = $Name.ToLowerInvariant()
    if ($script:KaFields[$TableId].ContainsKey($k)) { return $script:KaFields[$TableId][$k] }
    return $null
}

# ============================================================================
#  Разбор текста SQL: карта скобок и литералов
# ============================================================================

$script:KaIdent = '(?:"[^"]*"|\[[^\]]*\]|[A-Za-z_#@$][A-Za-z0-9_#@$]*)'

function Get-KaScan {
    <#
    .SYNOPSIS
        Карта текста запроса: глубина скобок и признак «внутри литерала» на каждый символ.

    .DESCRIPTION
        Нужна, чтобы делить WHERE на конъюнкты и искать ключевые слова, не попадаясь
        на подзапросы и на текст внутри кавычек. Сама скобка отнесена к внешнему уровню.
    #>
    param([string] $Sql)
    $n = $Sql.Length
    $depth = New-Object 'int[]'  $n
    $lit   = New-Object 'bool[]' $n
    $d = 0
    $mode = 0   # 0 код, 1 '…', 2 "…", 3 […]
    for ($i = 0; $i -lt $n; $i++) {
        $ch = $Sql[$i]
        if ($mode -eq 0) {
            if     ($ch -eq "'") { $mode = 1; $depth[$i] = $d; $lit[$i] = $true }
            elseif ($ch -eq '"') { $mode = 2; $depth[$i] = $d; $lit[$i] = $true }
            elseif ($ch -eq '[') { $mode = 3; $depth[$i] = $d; $lit[$i] = $true }
            elseif ($ch -eq '(') { $depth[$i] = $d; $d++ }
            elseif ($ch -eq ')') { if ($d -gt 0) { $d-- }; $depth[$i] = $d }
            else                 { $depth[$i] = $d }
        }
        else {
            $depth[$i] = $d; $lit[$i] = $true
            if     ($mode -eq 1 -and $ch -eq "'") { $mode = 0 }
            elseif ($mode -eq 2 -and $ch -eq '"') { $mode = 0 }
            elseif ($mode -eq 3 -and $ch -eq ']') { $mode = 0 }
        }
    }
    return [pscustomobject]@{ Sql = $Sql; Depth = $depth; Lit = $lit; Length = $n }
}

function Find-KaToken {
    <#
    .SYNOPSIS
        Позиции регулярки в запросе вне литералов, с фильтром по глубине скобок.
        -Depth -1 означает «любая глубина».
    #>
    param(
        [Parameter(Mandatory)] $Scan,
        [Parameter(Mandatory)] [string] $Pattern,
        [int] $Depth = -1
    )
    $res = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($Scan.Sql, $Pattern, 'IgnoreCase')) {
        $i = $m.Index
        if ($i -ge $Scan.Length) { continue }
        if ($Scan.Lit[$i]) { continue }
        if ($Depth -ge 0 -and $Scan.Depth[$i] -ne $Depth) { continue }
        [void]$res.Add($m)
    }
    return $res
}

function Compress-KaSpace {
    <#
    .SYNOPSIS
        Схлопывает пробелы и убирает комментарии, НЕ трогая литералы и кавычки.

    .DESCRIPTION
        Простая замена всех пробельных серий ломает реальные имена: в рабочей базе есть
        таблица «Production Matrix  BOM Entry» с ДВУМЯ пробелами, и после схлопывания
        она перестаёт находиться в словаре. Поэтому проход посимвольный, с учётом
        того, внутри кавычек мы или нет.
    #>
    param([string] $Sql)
    $sq = [char]39; $dq = [char]34; $ob = [char]91; $cb = [char]93
    $sb = New-Object System.Text.StringBuilder
    $n  = $Sql.Length
    $mode = 0        # 0 код, 1 апострофы, 2 двойные кавычки, 3 квадратные скобки
    $prevSpace = $false
    for ($i = 0; $i -lt $n; $i++) {
        $ch = $Sql[$i]
        if ($mode -ne 0) {
            [void]$sb.Append($ch); $prevSpace = $false
            if     ($mode -eq 1 -and $ch -eq $sq) { $mode = 0 }
            elseif ($mode -eq 2 -and $ch -eq $dq) { $mode = 0 }
            elseif ($mode -eq 3 -and $ch -eq $cb) { $mode = 0 }
            continue
        }
        # комментарии снимаются только вне литералов
        if ($ch -eq [char]45 -and $i + 1 -lt $n -and $Sql[$i + 1] -eq [char]45) {
            while ($i -lt $n -and $Sql[$i] -ne [char]10) { $i++ }
            if (-not $prevSpace) { [void]$sb.Append([char]32); $prevSpace = $true }
            continue
        }
        if ($ch -eq [char]47 -and $i + 1 -lt $n -and $Sql[$i + 1] -eq [char]42) {
            $i += 2
            while ($i + 1 -lt $n -and -not ($Sql[$i] -eq [char]42 -and $Sql[$i + 1] -eq [char]47)) { $i++ }
            $i++
            if (-not $prevSpace) { [void]$sb.Append([char]32); $prevSpace = $true }
            continue
        }
        if ([char]::IsWhiteSpace($ch)) {
            if (-not $prevSpace) { [void]$sb.Append([char]32); $prevSpace = $true }
            continue
        }
        $prevSpace = $false
        if     ($ch -eq $sq) { $mode = 1 }
        elseif ($ch -eq $dq) { $mode = 2 }
        elseif ($ch -eq $ob) { $mode = 3 }
        [void]$sb.Append($ch)
    }
    return $sb.ToString().Trim()
}

function Split-KaStatement {
    <#
    .SYNOPSIS
        Отделяет от текста запроса хвост NAV и подсказки оптимизатору.

    .DESCRIPTION
        NAV дописывает в конец «AppObjectType: Page AppObjectId: 21» — вызывающий
        объект: для профайлера ценная связка, для разбора SQL мусор. Ещё снимаются
        OPTION(...) и комментарии, пробелы схлопываются.
    #>
    param([string] $Sql)
    $callerType = ''; $callerId = ''
    if (-not $Sql) { return [pscustomobject]@{ Sql = ''; CallerType = ''; CallerId = '' } }

    $s = $Sql
    $m = [regex]::Match($s, 'AppObjectType:\s*(?<t>\w+)\s+AppObjectId:\s*(?<i>\d+)\s*$', 'IgnoreCase')
    if ($m.Success) {
        $callerType = $m.Groups['t'].Value
        $callerId   = $m.Groups['i'].Value
        $s = $s.Substring(0, $m.Index)
    }
    $s = Compress-KaSpace $s
    $s = [regex]::Replace($s, '\bOPTION\s*\([^()]*\)\s*$', ' ', 'IgnoreCase')
    $s = $s.Trim().TrimEnd([char]59).Trim()

    return [pscustomobject]@{ Sql = $s; CallerType = $callerType; CallerId = $callerId }
}

function Split-KaQuoted {
    <#
    .SYNOPSIS
        Снимает с идентификатора кавычки или квадратные скобки.
    #>
    param([string] $Ident)
    if (-not $Ident) { return '' }
    $s = $Ident.Trim()
    if ($s.Length -ge 2) {
        if     ($s[0] -eq '"' -and $s[$s.Length - 1] -eq '"') { return $s.Substring(1, $s.Length - 2) }
        elseif ($s[0] -eq '[' -and $s[$s.Length - 1] -eq ']') { return $s.Substring(1, $s.Length - 2) }
    }
    return $s
}

# ============================================================================
#  1. Get-SqlTableRef — какая таблица NAV стоит за именем в запросе
# ============================================================================

function Resolve-KaTableName {
    <#
    .SYNOPSIS
        SQL-имя объекта -> таблица NAV. При неоднозначности возвращает список
        кандидатов и НИЧЕГО не выбирает.
    #>
    param([string] $SqlName)
    Initialize-KeyAdvisor

    $raw = $SqlName
    $res = [pscustomobject]@{
        Status     = 'NotFound'
        SqlName    = $raw
        BaseName   = $raw
        Company    = ''
        TableId    = 0
        TableName  = ''
        IsSift     = $false
        SiftKeyNo  = 0
        IsSystem   = $false
        Candidates = @()
    }
    if (-not $raw) { return $res }

    # служебные таблицы платформы: $ndo$… и всё, что начинается с $
    if ($raw.StartsWith('$')) {
        $res.Status = 'System'; $res.IsSystem = $true
        return $res
    }

    $name = $raw

    # индексированное представление SIFT: <префикс>$<таблица>$VSIFT$<номер ключа - 1>
    $m = [regex]::Match($name, '^(?<base>.+)\$VSIFT\$(?<n>\d+)$', 'IgnoreCase')
    if ($m.Success) {
        $res.IsSift    = $true
        $res.SiftKeyNo = [int]$m.Groups['n'].Value + 1
        $name = $m.Groups['base'].Value
    }

    # сначала пробуем имя целиком: у таблиц с DataPerCompany=No префикса нет,
    # а имя таблицы NAV само может содержать $ — тогда отрезание префикса его испортит
    $tries = New-Object System.Collections.ArrayList
    [void]$tries.Add([pscustomobject]@{ Name = $name; Company = '' })
    $d = $name.IndexOf('$')
    if ($d -gt 0 -and $d -lt $name.Length - 1) {
        [void]$tries.Add([pscustomobject]@{ Name = $name.Substring($d + 1); Company = $name.Substring(0, $d) })
    }

    foreach ($t in $tries) {
        $k = $t.Name.ToLowerInvariant()
        if (-not $script:KaSqlMap.ContainsKey($k)) { continue }
        $cands = @($script:KaSqlMap[$k])
        $res.BaseName = $t.Name
        $res.Company  = $t.Company
        if ($cands.Count -gt 1) {
            $res.Status     = 'Ambiguous'
            $res.Candidates = $cands
            return $res
        }
        $res.Status    = 'Resolved'
        $res.TableId   = $cands[0].TableId
        $res.TableName = $cands[0].TableName
        return $res
    }

    # в словаре нет — сохраним разложение на префикс и имя для диагностики
    if ($tries.Count -gt 1) { $res.BaseName = $tries[1].Name; $res.Company = $tries[1].Company }
    return $res
}

function Get-SqlTableRef {
    <#
    .SYNOPSIS
        Из текста запроса — таблица NAV: номер, имя, признаки SIFT и служебной.

    .DESCRIPTION
        Разбираются реальные формы NAV и платформы:
            "NAV".dbo."NAV$Customer" "18" WITH(READUNCOMMITTED)
            [NAV].[dbo].[$ndo$taskscheduling]
            dbo.[NAV$G_L Entry]
            [Tenant Media]
            "NAV"."dbo"."NAV$Item Ledger Entry$VSIFT$3"
        FROM (SELECT …) и FROM CHANGETABLE(…) таблицей не считаются: вложенный запрос
        разбирается сам по себе, табличная функция пропускается.

        Псевдоним из одних цифр NAV делает равным НОМЕРУ таблицы — он служит
        независимой проверкой разбора имени и запасным путём, если имени нет
        в словаре. Расхождение имени и псевдонима — неоднозначность: решение
        не принимается.

    .PARAMETER Sql
        Текст запроса как есть, вместе с хвостом AppObjectType/AppObjectId.

    .OUTPUTS
        Ok, Status (Resolved|Ambiguous|NotFound|System|NoTable|NotAQuery), Reason, TableId,
        TableName, SqlName, Company, Alias, Source, IsSift, SiftKeyNo, IsSystem,
        Multi, Refs (все ссылки запроса), CallerType, CallerId.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Sql)

    Initialize-KeyAdvisor
    $st   = Split-KaStatement $Sql
    $text = $st.Sql
    $out  = [pscustomobject]@{
        Ok         = $false
        Status     = 'NoTable'
        Reason     = ''
        TableId    = 0
        TableName  = ''
        SqlName    = ''
        Company    = ''
        Alias      = ''
        Source     = ''
        IsSift     = $false
        SiftKeyNo  = 0
        IsSystem   = $false
        Multi      = $false
        Refs       = @()
        Candidates = @()
        CallerType = $st.CallerType
        CallerId   = $st.CallerId
    }
    if (-not $text) { $out.Reason = 'пустой запрос'; return $out }

    $scan = Get-KaScan $text
    $kw   = '\b(?:FROM|JOIN|UPDATE|INSERT\s+INTO|DELETE\s+FROM)\b'
    $hits = Find-KaToken -Scan $scan -Pattern $kw

    # слова, которые нельзя принять за псевдоним таблицы
    $notAlias = @('WITH','WHERE','ON','INNER','LEFT','RIGHT','FULL','OUTER','CROSS','JOIN',
                  'GROUP','ORDER','HAVING','UNION','OPTION','SET','VALUES','AS','FOR','WHEN',
                  'AND','OR','SELECT','INTO','EXCEPT','INTERSECT','INDEX')

    $refs = New-Object System.Collections.ArrayList
    foreach ($h in $hits) {
        $i = $h.Index + $h.Length
        while ($i -lt $text.Length -and $text[$i] -eq ' ') { $i++ }
        if ($i -ge $text.Length) { continue }
        if ($text[$i] -eq '(') { continue }   # подзапрос: его FROM найдётся отдельно

        # до трёх частей через точку: база, схема, объект
        $parts = New-Object System.Collections.Generic.List[string]
        $pos = $i
        while ($true) {
            $m = [regex]::Match($text.Substring($pos), "^$script:KaIdent")
            if (-not $m.Success) { break }
            $parts.Add($m.Value)
            $pos += $m.Length
            if ($pos -lt $text.Length -and $text[$pos] -eq '.') { $pos++; continue }
            break
        }
        if ($parts.Count -eq 0) { continue }
        # табличная функция: CHANGETABLE(…), OPENJSON(…) — сразу за именем скобка
        if ($pos -lt $text.Length -and $text[$pos] -eq '(') { continue }

        $objName = Split-KaQuoted $parts[$parts.Count - 1]

        # псевдоним: пробел, необязательное AS, идентификатор не из списка ключевых слов
        $alias = ''
        $ma = [regex]::Match($text.Substring($pos), "^\s+(?:AS\s+)?(?<a>$script:KaIdent)", 'IgnoreCase')
        if ($ma.Success) {
            $cand = Split-KaQuoted $ma.Groups['a'].Value
            if ($notAlias -notcontains $cand.ToUpperInvariant()) { $alias = $cand }
        }

        $r = Resolve-KaTableName $objName
        [void]$refs.Add([pscustomobject]@{
            Keyword    = ($h.Value -replace '\s+', ' ').ToUpperInvariant()
            Depth      = $scan.Depth[$h.Index]
            Position   = $h.Index
            SqlName    = $objName
            Alias      = $alias
            Status     = $r.Status
            TableId    = $r.TableId
            TableName  = $r.TableName
            Company    = $r.Company
            IsSift     = $r.IsSift
            SiftKeyNo  = $r.SiftKeyNo
            IsSystem   = $r.IsSystem
            Candidates = $r.Candidates
        })
    }

    $out.Refs = @($refs)
    if ($refs.Count -eq 0) {
        if ($hits.Count -eq 0) {
            $out.Status = 'NotAQuery'
            $out.Reason = 'обращения к таблице в тексте нет — это команда протокола, а не запрос'
        }
        else {
            $out.Reason = 'все источники запроса — подзапросы или табличные функции, базовой таблицы нет'
        }
        return $out
    }

    # основная — самая внешняя, при равной глубине первая по тексту
    $primary  = @($refs | Sort-Object Depth, Position)[0]
    $userRefs = @($refs | Where-Object { $_.Status -eq 'Resolved' } | Select-Object -ExpandProperty TableId -Unique)
    $out.Multi = ($userRefs.Count -gt 1)

    $out.SqlName    = $primary.SqlName
    $out.Company    = $primary.Company
    $out.Alias      = $primary.Alias
    $out.IsSift     = $primary.IsSift
    $out.SiftKeyNo  = $primary.SiftKeyNo
    $out.IsSystem   = $primary.IsSystem
    $out.Status     = $primary.Status
    $out.TableId    = $primary.TableId
    $out.TableName  = $primary.TableName
    $out.Candidates = $primary.Candidates

    # псевдоним из цифр = номер таблицы: независимая проверка разбора имени
    $aliasId = 0
    if ($primary.Alias -and $primary.Alias -match '^\d+$') {
        $aliasId = [int]$primary.Alias
        if (-not $script:KaTables.ContainsKey($aliasId)) { $aliasId = 0 }
    }

    switch ($primary.Status) {
        'Resolved' {
            $out.Ok = $true; $out.Source = 'имя'
            if ($aliasId -gt 0) {
                if ($aliasId -eq $primary.TableId) { $out.Source = 'имя+псевдоним' }
                else {
                    $out.Ok = $false; $out.Status = 'Ambiguous'
                    $out.Reason = ("имя даёт таблицу {0}, псевдоним — {1}" -f $primary.TableId, $aliasId)
                    $out.Candidates = @(
                        [pscustomobject]@{ TableId = $primary.TableId; TableName = $primary.TableName }
                        [pscustomobject]@{ TableId = $aliasId;         TableName = $script:KaTables[$aliasId] }
                    )
                }
            }
        }
        'Ambiguous' {
            $out.Reason = ("SQL-имя «{0}» соответствует нескольким таблицам" -f $primary.SqlName)
        }
        'System' {
            $out.Reason = 'служебная таблица платформы'
        }
        default {
            if ($aliasId -gt 0) {
                $out.Ok = $true; $out.Status = 'Resolved'; $out.Source = 'псевдоним'
                $out.TableId = $aliasId; $out.TableName = $script:KaTables[$aliasId]
                $out.Reason  = 'имени нет в словаре, таблица определена по номеру в псевдониме'
            }
            else {
                $out.Reason = ("SQL-имени «{0}» нет в словаре" -f $primary.SqlName)
            }
        }
    }
    return $out
}

# ============================================================================
#  2. Get-SqlPredicates — WHERE на множества E / R / N плюс ORDER BY
# ============================================================================

function Split-KaTerms {
    <#
    .SYNOPSIS
        Делит участок запроса на слагаемые по AND (или по OR) на заданной глубине скобок.

    .DESCRIPTION
        AND внутри BETWEEN … AND … точкой деления не является — иначе диапазон
        разваливается на два неразобранных куска.
    #>
    param(
        [Parameter(Mandatory)] $Scan,
        [Parameter(Mandatory)] [int] $Start,
        [Parameter(Mandatory)] [int] $End,
        [Parameter(Mandatory)] [int] $Depth,
        [ValidateSet('AND','OR')] [string] $Operator = 'AND'
    )
    $parts = New-Object System.Collections.Generic.List[string]
    $text  = $Scan.Sql
    $pend  = $false
    $from  = $Start
    foreach ($m in [regex]::Matches($text, '\b(?:BETWEEN|AND|OR)\b', 'IgnoreCase')) {
        if ($m.Index -lt $Start -or $m.Index -ge $End) { continue }
        if ($Scan.Lit[$m.Index]) { continue }
        if ($Scan.Depth[$m.Index] -ne $Depth) { continue }
        $w = $m.Value.ToUpperInvariant()
        if ($w -eq 'BETWEEN')        { $pend = $true; continue }
        if ($w -eq 'AND' -and $pend) { $pend = $false; continue }
        if ($w -ne $Operator)        { continue }
        $parts.Add($text.Substring($from, $m.Index - $from))
        $from = $m.Index + $m.Length
    }
    $parts.Add($text.Substring($from, $End - $from))
    return @($parts | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Resolve-KaWhereBounds {
    <#
    .SYNOPSIS
        Спускается внутрь объемлющих скобок WHERE: сдвигает границы и глубину.

    .DESCRIPTION
        NAV кладёт ВЕСЬ фильтр внутрь одной пары скобок:
        WHERE ("17"."No_"=@1 AND "17"."Name">=@2). Тогда все AND лежат на глубине
        на единицу больше, точек деления на нашей глубине нет вовсе, и весь WHERE
        уходит ОДНИМ слагаемым. Разобрать его как один предикат нельзя: жадная
        правая часть простого сравнения съедает остаток, и выходит единственный
        Unknown «правая часть — выражение», а множества E и R остаются пустыми.
        Совет при этом печатается — по пустому отбору, то есть неверный и молча.

        Спускаемся, оставаясь на той же карте: двигаются границы и глубина, текст
        не пересобирается, а значит и позиции из Get-KaScan остаются годными.

        Пара снимается, только если она ОБЪЕМЛЮЩАЯ. У «(A=1) AND (B=2)» первый и
        последний знаки тоже скобки, но внутри мы выныриваем на внешний уровень —
        такую пару трогать нельзя.
    #>
    param(
        [Parameter(Mandatory)] $Scan,
        [Parameter(Mandatory)] [int] $Start,
        [Parameter(Mandatory)] [int] $End,
        [Parameter(Mandatory)] [int] $Depth
    )
    $text = $Scan.Sql
    while ($true) {
        $a = $Start
        while ($a -lt $End -and [char]::IsWhiteSpace($text[$a])) { $a++ }
        $b = $End - 1
        while ($b -gt $a -and [char]::IsWhiteSpace($text[$b])) { $b-- }
        if ($a -ge $b) { break }
        if ($text[$a] -ne '(' -or $text[$b] -ne ')') { break }
        if ($Scan.Lit[$a] -or $Scan.Lit[$b]) { break }
        if ($Scan.Depth[$a] -ne $Depth -or $Scan.Depth[$b] -ne $Depth) { break }
        $escaped = $false
        for ($i = $a + 1; $i -lt $b; $i++) {
            if ($Scan.Depth[$i] -le $Depth) { $escaped = $true; break }
        }
        if ($escaped) { break }
        $Start = $a + 1
        $End   = $b
        $Depth = $Depth + 1
    }
    return [pscustomobject]@{ Start = $Start; End = $End; Depth = $Depth }
}

function Remove-KaOuterParens {
    param([string] $Term)
    $s = $Term.Trim()
    while ($s.Length -ge 2 -and $s[0] -eq '(' -and $s[$s.Length - 1] -eq ')') {
        $sc = Get-KaScan $s
        # Скобки лишние, только если закрывающая ДЛЯ ПЕРВОЙ — последний символ.
        #
        # Одной проверки «на последнем символе глубина 0» мало, и это не теория:
        # у «(A=1) OR (A=2)» она там тоже 0, потому что первая пара закрылась
        # раньше. Снятие рвало условие пополам — выходило «A=1) OR (A=2», разбор
        # такого куска давал Unknown, и вся группа по OR молча выпадала из отбора
        # индексируемых. Заходит сюда это с обычной формы «((A=1) OR (A=2))»:
        # первый проход снимает внешнюю пару правильно, а второй уже портит.
        #
        # Поэтому вторым условием: внутри мы ни разу не вынырнули на внешний
        # уровень. Скобка отнесена Get-KaScan к внешнему уровню, так что у своего
        # содержимого глубина не ниже 1.
        if ($sc.Depth[$s.Length - 1] -ne 0) { break }
        $escaped = $false
        for ($i = 1; $i -lt $s.Length - 1; $i++) {
            if ($sc.Depth[$i] -lt 1) { $escaped = $true; break }
        }
        if ($escaped) { break }
        $inner = $s.Substring(1, $s.Length - 2).Trim()
        if (-not $inner) { break }
        $s = $inner
    }
    return $s
}

function Get-KaColumnName {
    <#
    .SYNOPSIS
        Из «"17"."Posting Date"» — квалификатор и имя колонки.
    #>
    param([string] $Ref)
    $m = [regex]::Match($Ref.Trim(), "^(?:(?<q>$script:KaIdent)\s*\.\s*)?(?<c>$script:KaIdent)$")
    if (-not $m.Success) { return $null }
    return [pscustomobject]@{
        Qualifier = (Split-KaQuoted $m.Groups['q'].Value)
        Column    = (Split-KaQuoted $m.Groups['c'].Value)
    }
}

function New-KaPredicate {
    param([string]$Class, [string]$Column, [string]$Qualifier, [string]$Op, [string]$Reason, [string]$Term)
    return [pscustomobject]@{
        Class     = $Class      # Equality | Range | NonIndexable | Join | Unknown
        Column    = $Column
        Qualifier = $Qualifier
        Op        = $Op
        Reason    = $Reason
        Term      = $Term
    }
}

function Get-KaOperandKind {
    <#
    .SYNOPSIS
        Что стоит справа от оператора: параметр, литерал, колонка, выражение.
    #>
    param([string] $Operand)
    $s = $Operand.Trim()
    if (-not $s)                                    { return 'Пусто' }
    if ($s -match '^@\w+$')                         { return 'Параметр' }
    if ($s -match '^N?\x27')                        { return 'Литерал' }
    if ($s -match '^[-+]?\d')                       { return 'Литерал' }
    if ($s -match '^0x')                            { return 'Литерал' }
    if ($s -match '^\{\s*(ts|d|t)\b')               { return 'Литерал' }
    if ($s -match '^(?i:NULL)$')                    { return 'Литерал' }
    if ($s -match '^(?i:GET(UTC)?DATE\s*\(\s*\))$') { return 'Литерал' }
    if (Get-KaColumnName $s)                        { return 'Колонка' }
    return 'Выражение'
}

function Get-KaLikeShape {
    <#
    .SYNOPSIS
        Форма шаблона LIKE: префикс (индексируемо), маска (нет) или скрыт параметром.

    .DESCRIPTION
        NAV почти всегда прячет шаблон в параметр, и тогда префикс от маски не отличить.
        Такой предикат уходит в Unknown, а не в индексируемые: осторожность важнее.
    #>
    param([string] $Operand)
    $s = $Operand.Trim()
    if ($s -match '^@\w+$') { return 'Параметр' }
    $m = [regex]::Match($s, '^N?\x27(?<p>.*)\x27$', 'Singleline')
    if (-not $m.Success) { return 'Неизвестно' }
    $p = $m.Groups['p'].Value
    if ($p.Length -eq 0) { return 'Маска' }
    if ($p[0] -eq '%' -or $p[0] -eq '_' -or $p[0] -eq '[') { return 'Маска' }
    return 'Префикс'
}

function Get-KaTermPredicate {
    <#
    .SYNOPSIS
        Разбирает один конъюнкт WHERE в предикат нужного класса.
    #>
    param([string] $Term)

    $t = Remove-KaOuterParens $Term
    if (-not $t) { return New-KaPredicate 'Unknown' '' '' '' 'пустое условие' $Term }
    $col1 = "(?<col>(?:$script:KaIdent\s*\.\s*)?$script:KaIdent)"

    # группа по OR: индексируема, только если все ветви — по одному и тому же полю
    $sc  = Get-KaScan $t
    $ors = @(Find-KaToken -Scan $sc -Pattern '\bOR\b' -Depth 0)
    if ($ors.Count -gt 0) {
        $subs  = Split-KaTerms -Scan $sc -Start 0 -End $t.Length -Depth 0 -Operator 'OR'
        $inner = @($subs | ForEach-Object { Get-KaTermPredicate $_ })
        $cols  = @($inner | Where-Object { $_.Column } | Select-Object -ExpandProperty Column -Unique)
        $bad   = @($inner | Where-Object { $_.Class -eq 'NonIndexable' -or $_.Class -eq 'Unknown' })
        if ($cols.Count -eq 1 -and $bad.Count -eq 0) {
            $allEq = (@($inner | Where-Object { $_.Class -ne 'Equality' }).Count -eq 0)
            if ($allEq) { return New-KaPredicate 'Equality' $cols[0] $inner[0].Qualifier 'OR' 'набор значений одного поля' $Term }
            return New-KaPredicate 'Range' $cols[0] $inner[0].Qualifier 'OR' 'набор диапазонов одного поля' $Term
        }
        $why = 'OR по разным полям'
        if ($cols.Count -le 1) { $why = 'OR с неиндексируемой ветвью' }
        return New-KaPredicate 'NonIndexable' '' '' 'OR' $why $Term
    }

    if ($t -match '^(?i:NOT)\s') {
        return New-KaPredicate 'NonIndexable' '' '' 'NOT' 'отрицание условия' $Term
    }

    # IS [NOT] NULL
    $m = [regex]::Match($t, "^$col1\s+IS\s+(?<not>NOT\s+)?NULL$", 'IgnoreCase')
    if ($m.Success) {
        $c = Get-KaColumnName $m.Groups['col'].Value
        if ($m.Groups['not'].Success) {
            return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier 'IS NOT NULL' 'отрицание' $Term
        }
        return New-KaPredicate 'Equality' $c.Column $c.Qualifier 'IS NULL' '' $Term
    }

    # BETWEEN
    $m = [regex]::Match($t, "^$col1\s+(?<not>NOT\s+)?BETWEEN\s+", 'IgnoreCase')
    if ($m.Success) {
        $c = Get-KaColumnName $m.Groups['col'].Value
        if ($m.Groups['not'].Success) {
            return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier 'NOT BETWEEN' 'отрицание' $Term
        }
        return New-KaPredicate 'Range' $c.Column $c.Qualifier 'BETWEEN' '' $Term
    }

    # IN (…)
    $m = [regex]::Match($t, "^$col1\s+(?<not>NOT\s+)?IN\s*\((?<list>.*)\)$", 'IgnoreCase')
    if ($m.Success) {
        $c = Get-KaColumnName $m.Groups['col'].Value
        if ($m.Groups['not'].Success) {
            return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier 'NOT IN' 'отрицание' $Term
        }
        if ($m.Groups['list'].Value -match '(?i)\bSELECT\b') {
            return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier 'IN (SELECT)' 'подзапрос' $Term
        }
        return New-KaPredicate 'Equality' $c.Column $c.Qualifier 'IN' 'набор значений' $Term
    }

    # LIKE
    $m = [regex]::Match($t, "^$col1\s+(?<not>NOT\s+)?LIKE\s+(?<rhs>.+?)(?:\s+ESCAPE\s+.+)?$", 'IgnoreCase')
    if ($m.Success) {
        $c = Get-KaColumnName $m.Groups['col'].Value
        if ($m.Groups['not'].Success) {
            return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier 'NOT LIKE' 'отрицание' $Term
        }
        switch (Get-KaLikeShape $m.Groups['rhs'].Value) {
            'Префикс'  { return New-KaPredicate 'Range'        $c.Column $c.Qualifier 'LIKE' 'шаблон-префикс' $Term }
            'Параметр' { return New-KaPredicate 'Unknown'      $c.Column $c.Qualifier 'LIKE' 'шаблон скрыт параметром — префикс от маски не отличить' $Term }
            default    { return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier 'LIKE' 'шаблон с маской слева' $Term }
        }
    }

    # функция над колонкой: индекс по этой колонке при таком условии не работает
    $m = [regex]::Match($t, '^(?<fn>[A-Za-z_]\w*)\s*\(', 'IgnoreCase')
    if ($m.Success) {
        $col = ''
        $mc = [regex]::Match($t, "(?<c>$script:KaIdent)\s*[,)]")
        if ($mc.Success) { $col = Split-KaQuoted $mc.Groups['c'].Value }
        return New-KaPredicate 'NonIndexable' $col '' $m.Groups['fn'].Value.ToUpperInvariant() 'функция над колонкой' $Term
    }

    # простое сравнение
    $m = [regex]::Match($t, "^$col1\s*(?<op><>|!=|>=|<=|=|>|<)\s*(?<rhs>.+)$")
    if ($m.Success) {
        $c    = Get-KaColumnName $m.Groups['col'].Value
        $op   = $m.Groups['op'].Value
        $kind = Get-KaOperandKind $m.Groups['rhs'].Value.Trim()
        if ($kind -eq 'Колонка') {
            return New-KaPredicate 'Join' $c.Column $c.Qualifier $op 'сравнение двух колонок' $Term
        }
        if ($kind -eq 'Выражение') {
            return New-KaPredicate 'Unknown' $c.Column $c.Qualifier $op 'правая часть — выражение' $Term
        }
        switch ($op) {
            '='     { return New-KaPredicate 'Equality'     $c.Column $c.Qualifier '='  '' $Term }
            '<>'    { return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier '<>' 'неравенство' $Term }
            '!='    { return New-KaPredicate 'NonIndexable' $c.Column $c.Qualifier '<>' 'неравенство' $Term }
            default { return New-KaPredicate 'Range'        $c.Column $c.Qualifier $op  '' $Term }
        }
    }

    return New-KaPredicate 'Unknown' '' '' '' 'условие не разобрано' $Term
}

function Get-SqlPredicates {
    <#
    .SYNOPSIS
        Разбирает WHERE на три множества и вытаскивает ORDER BY и агрегаты.

    .DESCRIPTION
        E — равенство: =@n, IN (…), IS NULL, а также группа OR по ОДНОМУ полю.
        R — диапазон: >=, <=, >, <, BETWEEN, LIKE с шаблоном-префиксом.
        N — не индексируется: <>, NOT …, LIKE с маской слева, функция над колонкой,
            OR по разным полям, IN (SELECT …).
        Отдельно копится Unknown — то, чего разбор не понял; туда же уходит LIKE @n,
        потому что NAV прячет шаблон в параметр и префикс от маски не отличить.
        Поля из N и Unknown в предлагаемый ключ не попадают никогда.

        Направление сортировки игнорируется: индекс читается в обе стороны.

    .PARAMETER Sql
        Текст запроса.

    .PARAMETER Qualifier
        Если задан, берутся только предикаты по колонкам этого псевдонима
        (и неквалифицированные). Нужно для запросов с JOIN.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Sql,
        [string] $Qualifier
    )
    Initialize-KeyAdvisor
    $st   = Split-KaStatement $Sql
    $text = $st.Sql

    $out = [pscustomobject]@{
        Ok             = $false
        HasWhere       = $false
        WhereFromInner = $false
        WhereText      = ''
        Equality       = @()
        Range          = @()
        NonIndexable   = @()
        Unknown        = @()
        Joins          = @()
        Predicates     = @()
        OrderBy        = @()
        Aggregates     = @()
        Statement      = 'Прочее'
        HasTop         = $false
        HasDistinct    = $false
        Reason         = ''
    }
    if (-not $text) { $out.Reason = 'пустой запрос'; return $out }

    $scan = Get-KaScan $text
    if ($text -match '^\s*(?<w>SELECT|UPDATE|DELETE|INSERT|MERGE)\b') { $out.Statement = $Matches['w'].ToUpperInvariant() }
    $out.HasTop      = [regex]::IsMatch($text, '\bTOP\s*[\(\d@]', 'IgnoreCase')
    $out.HasDistinct = [regex]::IsMatch($text, '^\s*SELECT\s+DISTINCT\b', 'IgnoreCase')

    # --- агрегаты -----------------------------------------------------------
    $aggs = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches($text, '\b(?<fn>SUM|COUNT|MIN|MAX|AVG)\s*\(\s*(?<arg>[^()]*)\)', 'IgnoreCase')) {
        if ($scan.Lit[$m.Index]) { continue }
        $arg = [regex]::Replace($m.Groups['arg'].Value.Trim(), '^(?i:DISTINCT)\s+', '')
        $col = ''
        if ($arg -ne '*') {
            $c = Get-KaColumnName $arg
            if ($c) { $col = $c.Column }
        }
        $field = $col
        # в представлении SIFT колонка суммы называется SUM$<поле>
        $mf = [regex]::Match($field, '^SUM\$(?<f>.+)$')
        if ($mf.Success) { $field = $mf.Groups['f'].Value }
        [void]$aggs.Add([pscustomobject]@{
            Function = $m.Groups['fn'].Value.ToUpperInvariant()
            Argument = $arg
            Column   = $col
            Field    = $field
        })
    }
    $out.Aggregates = @($aggs)

    # --- ORDER BY -----------------------------------------------------------
    $ob = @(Find-KaToken -Scan $scan -Pattern '\bORDER\s+BY\b')
    if ($ob.Count -gt 0) {
        $obm   = $ob[0]
        $depth = $scan.Depth[$obm.Index]
        $start = $obm.Index + $obm.Length
        $end   = $text.Length
        foreach ($m in [regex]::Matches($text, '\b(?:OPTION|FOR|UNION|EXCEPT|INTERSECT)\b|\)', 'IgnoreCase')) {
            if ($m.Index -le $start) { continue }
            if ($scan.Lit[$m.Index]) { continue }
            if ($m.Value -eq ')') {
                if ($scan.Depth[$m.Index] -lt $depth) { $end = $m.Index; break }
                continue
            }
            if ($scan.Depth[$m.Index] -ne $depth) { continue }
            $end = $m.Index; break
        }
        # ORDER BY делится запятыми на своей глубине
        $seg     = $text.Substring($start, $end - $start)
        $segScan = Get-KaScan $seg
        $list    = New-Object System.Collections.Generic.List[string]
        $from    = 0
        for ($i = 0; $i -lt $seg.Length; $i++) {
            if ($seg[$i] -eq ',' -and -not $segScan.Lit[$i] -and $segScan.Depth[$i] -eq 0) {
                $list.Add($seg.Substring($from, $i - $from)); $from = $i + 1
            }
        }
        $list.Add($seg.Substring($from))
        $cols = New-Object System.Collections.Generic.List[string]
        foreach ($it in $list) {
            $s = [regex]::Replace($it.Trim(), '\s+(?i:ASC|DESC)$', '').Trim()
            if (-not $s) { continue }
            $c = Get-KaColumnName $s
            if ($c) { $cols.Add($c.Column) }
        }
        $out.OrderBy = @($cols)
    }

    # --- WHERE --------------------------------------------------------------
    $wh = @(Find-KaToken -Scan $scan -Pattern '\bWHERE\b')
    if ($wh.Count -eq 0) { $out.Ok = $true; $out.Reason = 'условий нет'; return $out }

    $wm    = @($wh | Sort-Object { $scan.Depth[$_.Index] }, Index)[0]
    $depth = $scan.Depth[$wm.Index]
    $out.HasWhere       = $true
    $out.WhereFromInner = ($depth -gt 0)

    $start = $wm.Index + $wm.Length
    $end   = $text.Length
    foreach ($m in [regex]::Matches($text, '\b(?:GROUP\s+BY|ORDER\s+BY|HAVING|OPTION|UNION|EXCEPT|INTERSECT|FOR\s+XML)\b|\)', 'IgnoreCase')) {
        if ($m.Index -le $start) { continue }
        if ($scan.Lit[$m.Index]) { continue }
        if ($m.Value -eq ')') {
            if ($scan.Depth[$m.Index] -lt $depth) { $end = $m.Index; break }
            continue
        }
        if ($scan.Depth[$m.Index] -ne $depth) { continue }
        $end = $m.Index; break
    }
    $out.WhereText = $text.Substring($start, $end - $start).Trim()

    # Внешние скобки всего фильтра - форма самого NAV; без спуска внутрь делить
    # нечего и весь WHERE ушёл бы одним неразобранным слагаемым.
    $wb    = Resolve-KaWhereBounds -Scan $scan -Start $start -End $end -Depth $depth
    $terms = Split-KaTerms -Scan $scan -Start $wb.Start -End $wb.End -Depth $wb.Depth -Operator 'AND'
    $preds = New-Object System.Collections.ArrayList
    foreach ($t in $terms) {
        $p = Get-KaTermPredicate $t
        if ($Qualifier -and $p.Qualifier -and $p.Qualifier -ne $Qualifier) { continue }
        [void]$preds.Add($p)
    }
    $out.Predicates   = @($preds)
    $out.Equality     = @($preds | Where-Object { $_.Class -eq 'Equality' })
    $out.Range        = @($preds | Where-Object { $_.Class -eq 'Range' })
    $out.NonIndexable = @($preds | Where-Object { $_.Class -eq 'NonIndexable' })
    $out.Unknown      = @($preds | Where-Object { $_.Class -eq 'Unknown' })
    $out.Joins        = @($preds | Where-Object { $_.Class -eq 'Join' })
    $out.Ok           = $true
    return $out
}

# ============================================================================
#  3. Test-KeyCoverage — что из существующих ключей ложится на запрос
# ============================================================================

function Test-KeyCoverage {
    <#
    .SYNOPSIS
        Сопоставляет множества E / R / ORDER BY с ключами таблицы.

    .DESCRIPTION
        Правило: множество E должно совпасть с непрерывным ПРЕФИКСОМ ключа — порядок
        полей внутри E не важен, потому что все они сравниваются на равенство. Поле
        диапазона R обязано стоять ровно на позиции |E|+1: дальше по ключу просмотр
        уже не сходится в один диапазон. Сортировка считается покрытой, если после
        выбрасывания полей из E остаток ORDER BY идёт подряд с позиции |E|.
        Удлинение предлагается только тогда, когда ключ ЦЕЛИКОМ лежит внутри E: если
        в ключе есть поле, которого в фильтре нет, дописывание в конец префикс
        не восстановит.

        Сравнение ведётся по составу ИНДЕКСА (свойство SQLIndex перебивает состав
        ключа), а совет даётся по составу КЛЮЧА — SETCURRENTKEY знает только его.

    .PARAMETER TableId
        Номер таблицы NAV.

    .PARAMETER Equality
        Поля из множества E.

    .PARAMETER Range
        Поля из множества R (в индекс идёт первое, остальные — в замечания).

    .PARAMETER OrderBy
        Поля сортировки; направление не важно.

    .PARAMETER Company
        Префикс компании из запроса. Нужен только чтобы имя индекса совпало с тем,
        что видно в sys.indexes: у таблицы с DataPerCompany=Yes кластерный индекс
        называется «NAV$G_L Entry$0», а не «G_L Entry$0».

    .PARAMETER HasOtherPredicates
        Признак «условия были, но все неиндексируемые». Меняет вердикт при пустых
        E/R/ORDER BY с NoPredicates на NoIndexablePredicate: полный перебор в обоих
        случаях, но причина и лечение разные.

    .OUTPUTS
        Level (Full|Seek|Partial|Sort|None), Verdict (KeyExists|KeyDisabled|
        KeyNotMaintained|ExtendKey|NewKey|NoIndexablePredicate|NoPredicates|NoTable),
        BestKey, BestKeyNo,
        SqlIndexName, MatchedPrefix, OrderBySatisfied, RangeAtNextPos, Notes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int] $TableId,
        [string[]] $Equality = @(),
        [string[]] $Range    = @(),
        [string[]] $OrderBy  = @(),
        [string]   $Company,
        [switch]   $HasOtherPredicates
    )
    Initialize-KeyAdvisor

    $tblName = ''
    if ($script:KaTables.ContainsKey($TableId)) { $tblName = $script:KaTables[$TableId] }

    $res = [pscustomobject]@{
        TableId          = $TableId
        TableName        = $tblName
        Level            = 'None'
        Verdict          = 'NewKey'
        BestKey          = $null
        BestKeyNo        = 0
        SqlIndexName     = ''
        MatchedPrefix    = 0
        OrderBySatisfied = $false
        RangeAtNextPos   = $false
        Equality         = @($Equality)
        Range            = @($Range)
        OrderBy          = @($OrderBy)
        Candidates       = @()
        Notes            = @()
    }

    $keys = Get-KaTableKeys $TableId
    if (@($keys).Count -eq 0) {
        $res.Verdict = 'NoTable'
        $res.Notes  += ("по таблице {0} ключей в справочнике нет" -f $TableId)
        return $res
    }

    # имена приводим к именам полей C/AL: они пойдут в текст совета
    $Equality = @($Equality | ForEach-Object { ConvertTo-KaFieldName -TableId $TableId -Name $_ })
    $Range    = @($Range    | ForEach-Object { ConvertTo-KaFieldName -TableId $TableId -Name $_ })
    $OrderBy  = @($OrderBy  | ForEach-Object { ConvertTo-KaFieldName -TableId $TableId -Name $_ })

    # FlowField и FlowFilter своей колонки в SQL не имеют и в индекс попасть не могут;
    # выключенное поле - тоже. New-KeyProposal их отбрасывает, а покрытие считало
    # наравне с обычными: выходил ExtendKey с пустым списком того, что дописывать,
    # там где ключ на самом деле есть. Отбрасываем ЗДЕСЬ же, тем же правилом.
    $dropped = New-Object System.Collections.Generic.List[string]
    function Test-KaIndexable([int] $Tid, [string] $Name, $Bag) {
        if (-not $Name) { return $false }
        $fld = Get-KaField -TableId $Tid -Name $Name
        if ($fld -and ($fld.FieldClass -ne 'Normal' -or -not $fld.Enabled)) {
            [void]$Bag.Add($Name)
            return $false
        }
        return $true
    }
    $Equality = @($Equality | Where-Object { Test-KaIndexable $TableId $_ $dropped })
    $Range    = @($Range    | Where-Object { Test-KaIndexable $TableId $_ $dropped })
    $OrderBy  = @($OrderBy  | Where-Object { Test-KaIndexable $TableId $_ $dropped })
    foreach ($d in @($dropped | Select-Object -Unique)) {
        $res.Notes += ("{0} — своей колонки в SQL нет либо поле выключено, в покрытии не учитывается" -f $d)
    }
    $res.Equality = @($Equality); $res.Range = @($Range); $res.OrderBy = @($OrderBy)

    $eL  = @($Equality | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() } | Select-Object -Unique)
    $rL  = @($Range    | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() } | Select-Object -Unique)
    $obL = @($OrderBy  | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() })
    $r1  = ''
    if ($rL.Count -gt 0) { $r1 = $rL[0] }

    if ($eL.Count -eq 0 -and $rL.Count -eq 0 -and $obL.Count -eq 0) {
        if ($HasOtherPredicates) {
            $res.Verdict = 'NoIndexablePredicate'
            $res.Notes  += 'условия в запросе есть, но ни одно не ложится в индекс — читается вся таблица; лечится не ключом, а переписыванием фильтра'
        }
        else {
            $res.Verdict = 'NoPredicates'
            $res.Notes  += 'ни фильтров, ни сортировки — индекс тут ни при чём, это полный перебор'
        }
        return $res
    }

    $scored = New-Object System.Collections.ArrayList
    foreach ($k in $keys) {
        $kf = $k.EffectiveLower
        $m  = 0
        while ($m -lt $kf.Count -and ($eL -contains $kf[$m])) { $m++ }
        # При ПУСТОМ множестве равенств покрывать нечего, и прежнее ($m -eq 0) было
        # истиной у ЛЮБОГО ключа: фильтр с одним диапазоном получал «ключ есть» по
        # первому попавшемуся ключу, и настоящее «ключа нет» пропадало молча.
        # Поиск по индексу при пустом E возможен, но только когда диапазон стоит
        # ПЕРВЫМ полем ключа, - это отдельный случай ниже.
        $eCovered = ($eL.Count -gt 0 -and $m -eq $eL.Count)

        $rOk = $true
        if ($r1) { $rOk = ($m -lt $kf.Count -and $kf[$m] -eq $r1) }

        $obOk = $true
        $j = $m
        foreach ($o in $obL) {
            if ($eL -contains $o) { continue }
            if ($j -lt $kf.Count -and $kf[$j] -eq $o) { $j++ } else { $obOk = $false; break }
        }

        # Удлинять имеет смысл, только когда ключ ЦЕЛИКОМ лежит внутри E: если в
        # ключе есть поле, которого в фильтре нет, дописывание в конец не поможет —
        # префикс всё равно оборвётся на этом поле.
        $insideE = ($m -gt 0 -and $m -eq @($k.FieldsLower).Count)
        # Пустое E, но диапазон стоит первым полем ключа - это тоже поиск по индексу,
        # просто по диапазону, а не по равенству.
        $seekByRange = ($eL.Count -eq 0 -and $r1 -and $rOk)
        $level = 'None'
        if     ($eCovered -and $rOk -and $obOk) { $level = 'Full' }
        elseif ($seekByRange -and $obOk)        { $level = 'Full' }
        elseif ($eCovered -or $seekByRange)     { $level = 'Seek' }
        elseif ($insideE)                       { $level = 'Partial' }
        if ($eL.Count -eq 0 -and $rL.Count -eq 0 -and $obL.Count -gt 0 -and $obOk) { $level = 'Sort' }

        $score = 0
        if ($eCovered -or $seekByRange)           { $score += 1000 }
        if (($eCovered -or $seekByRange) -and $rOk) { $score += 200 }
        if ($obOk)               { $score += 100 }
        if ($level -eq 'Sort')   { $score += 400 }
        $score += $m * 5
        if ($k.Enabled)          { $score += 30 }
        if ($k.MaintainSQLIndex) { $score += 20 }
        $score -= $kf.Count      # при равном покрытии узкий индекс лучше

        [void]$scored.Add([pscustomobject]@{
            Key              = $k
            Level            = $level
            Score            = $score
            MatchedPrefix    = $m
            RangeAtNextPos   = $rOk
            OrderBySatisfied = $obOk
        })
    }

    $ordered = @($scored | Sort-Object -Property @{Expression='Score';Descending=$true}, @{Expression={$_.Key.KeyNo};Descending=$false})
    $best    = $ordered[0]
    $res.Candidates       = @($ordered | Select-Object -First 3)
    $res.Level            = $best.Level
    $res.BestKey          = $best.Key
    $res.BestKeyNo        = $best.Key.KeyNo
    $res.MatchedPrefix    = $best.MatchedPrefix
    $res.RangeAtNextPos   = $best.RangeAtNextPos
    $res.OrderBySatisfied = $best.OrderBySatisfied

    # имя индекса в SQL: ключ 1 — <префикс компании>$<SQL-имя таблицы>$0, остальные — $(N-1)
    if ($best.Key.KeyNo -eq 1) {
        $pfx = ''
        if ($Company) { $pfx = $Company + '$' }
        $res.SqlIndexName = ('{0}{1}$0' -f $pfx, (ConvertTo-KaSqlName $best.Key.TableName))
    }
    else { $res.SqlIndexName = $best.Key.IndexSuffix }

    switch ($best.Level) {
        'Full'    { $res.Verdict = 'KeyExists' }
        'Sort'    { $res.Verdict = 'KeyExists' }
        'Seek'    { $res.Verdict = 'KeyExists' }
        'Partial' { $res.Verdict = 'ExtendKey' }
        default   { $res.Verdict = 'NewKey' }
    }

    if ($res.Verdict -eq 'KeyExists') {
        if (-not $best.Key.Enabled) {
            $res.Verdict = 'KeyDisabled'
            $res.Notes  += ("ключ {0} выключен (Enabled=No) — индекса {1} в SQL нет; нумерация ключей сквозная, включение номера остальным не сдвинет" -f $best.Key.KeyNo, $res.SqlIndexName)
        }
        elseif (-not $best.Key.MaintainSQLIndex) {
            $res.Verdict = 'KeyNotMaintained'
            $res.Notes  += ("ключ {0} есть, но MaintainSQLIndex=No — индекса в SQL нет, ключ работает только как порядок обхода" -f $best.Key.KeyNo)
        }
    }
    if ($best.Level -eq 'Seek' -and $obL.Count -gt 0 -and -not $best.OrderBySatisfied) {
        $res.Notes += 'фильтр на ключ ложится, сортировка — нет: SQL добавит Sort по всей выборке'
    }
    if ($best.Level -eq 'Seek' -and $r1 -and -not $best.RangeAtNextPos) {
        $res.Notes += ("поле диапазона «{0}» стоит в ключе не сразу за полями равенства — часть строк отсеется уже после чтения индекса" -f $Range[0])
    }
    if (@($best.Key.SqlIndexFields).Count -gt 0) {
        $res.Notes += ('у ключа задан SQLIndex — состав индекса в SQL отличается от состава ключа: ' + (@($best.Key.SqlIndexFields) -join ', '))
    }
    if ($rL.Count -gt 1) {
        $res.Notes += ('полей диапазона больше одного — в индекс встанет только первое: ' + (@($Range) -join ', '))
    }
    return $res
}

# ============================================================================
#  5. Get-TableClass — что вообще можно советовать по этой таблице
# ============================================================================

function Get-TableClass {
    <#
    .SYNOPSIS
        Класс таблицы и политика совета по нему.

    .DESCRIPTION
        Custom     50000–99999             — клиентский диапазон, совет полноценный;
        Standard   1–49999, 100000–9999999 — стандартные MS и локализация: совет
                   с предупреждением про upgrade-контур и требованием сперва искать
                   выключенный ключ нужного состава;
        ClientCore 10000000+               — ядро отраслевого решения (ЕГАИС, маркировка,
                   фискалка, OMS): правка требует лицензии владельца решения, поэтому
                   только предупреждение, без «сделай»;
        Virtual    2000000000+ и служебные $ndo$ — молчим, ключей у них нет.
    #>
    [CmdletBinding()]
    param(
        [int]    $TableId,
        [string] $SqlName
    )
    if ($SqlName -and $SqlName.StartsWith('$')) {
        return [pscustomobject]@{
            Class  = 'System'; Advise = $false; Silent = $true
            Title  = 'служебная таблица платформы'
            Policy = 'Ключи задаёт платформа, вмешательство недопустимо.'
        }
    }
    if ($TableId -le 0) {
        return [pscustomobject]@{
            Class  = 'Unknown'; Advise = $false; Silent = $true
            Title  = 'таблица не определена'
            Policy = 'Без номера таблицы совет не даётся.'
        }
    }
    if ($TableId -ge $script:KaRangeVirtualFrom) {
        return [pscustomobject]@{
            Class  = 'Virtual'; Advise = $false; Silent = $true
            Title  = 'виртуальная или системная таблица'
            Policy = 'Хранилища и ключей в обычном смысле нет — советовать нечего.'
        }
    }
    if ($TableId -ge $script:KaRangeClientCoreFrom) {
        return [pscustomobject]@{
            Class  = 'ClientCore'; Advise = $false; Silent = $false
            Title  = 'ядро отраслевого решения (диапазон 10 000 000+)'
            Policy = 'Только сигнал: правка объектов этого диапазона требует лицензии владельца решения. Ключ заводит владелец подсистемы, дело профайлера — передать ему замер.'
        }
    }
    if ($TableId -ge $script:KaRangeCustomFrom -and $TableId -le $script:KaRangeCustomTo) {
        return [pscustomobject]@{
            Class  = 'Custom'; Advise = $true; Silent = $false
            Title  = 'клиентский диапазон 50000–99999'
            Policy = 'Ключ можно завести своей доработкой.'
        }
    }
    return [pscustomobject]@{
        Class  = 'Standard'; Advise = $true; Silent = $false
        Title  = 'стандартная таблица (MS или локализация)'
        Policy = 'Правка уходит в upgrade-контур: при обновлении всплывёт как конфликт. Сначала искать среди ВЫКЛЮЧЕННЫХ ключей готовый нужного состава — его включение конфликта не создаёт.'
    }
}

# ============================================================================
#  6. Test-SiftCoverage — агрегат мимо SIFT
# ============================================================================

function Test-SiftCoverage {
    <#
    .SYNOPSIS
        Считается ли агрегат перебором строк вместо индексированного представления SIFT.

    .DESCRIPTION
        Сигнал точный, не эвристический: если запрос считает SUM по БАЗОВОЙ таблице,
        а не по представлению «…$VSIFT$n», значит SIFT в этом месте не работает.
        Состав нужного ключа берётся прямо из запроса — поля равенства, — а поле
        агрегата из самого SUM(). Дальше сверяется, нет ли уже ключа с таким
        SumIndexFields: тогда причина не в отсутствии ключа, а в том, что фильтр
        не лёг на его префикс, либо у ключа MaintainSIFTIndex=No и представления
        просто нет.

    .PARAMETER TableRef
        Результат Get-SqlTableRef.

    .PARAMETER Predicates
        Результат Get-SqlPredicates.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $TableRef,
        [Parameter(Mandatory)] $Predicates
    )
    Initialize-KeyAdvisor

    $out = [pscustomobject]@{
        IsAggregate     = $false
        ViaSift         = $false
        Function        = ''
        Field           = ''
        NeededKeyFields = @()
        ExistingKeyNo   = 0
        ExistingKey     = $null
        SiftViewName    = ''
        Verdict         = 'NotAggregate'
        Note            = ''
    }
    if (@($Predicates.Aggregates).Count -eq 0) { return $out }

    $out.IsAggregate = $true
    $out.ViaSift     = [bool]$TableRef.IsSift
    # COUNT(*) агрегатом по полю не считаем: под него SumIndexFields не заводят
    $byField = @($Predicates.Aggregates | Where-Object { $_.Field })
    $a = @($Predicates.Aggregates)[0]
    if (@($byField).Count -gt 0) { $a = $byField[0] }
    $out.Function = $a.Function
    $out.Field    = $a.Field

    if ($TableRef.IsSift) {
        $out.Verdict = 'ViaSift'
        $out.Note    = ("агрегат идёт через представление SIFT ключа {0} — это и есть штатный путь" -f $TableRef.SiftKeyNo)
        return $out
    }
    if (-not $TableRef.Ok -or $TableRef.TableId -le 0) {
        $out.Verdict = 'Unresolved'
        $out.Note    = 'таблица не определена, судить о SIFT нельзя'
        return $out
    }

    # COUNT(*) считает СТРОКИ, а не сумму поля: SumIndexFields под него не заводят,
    # и SIFT тут ни при чём. Комментарий выше это и говорил, а кода не было - совет
    # печатался про сумму «» с пустым именем поля, потому что $byField пуст и
    # первым берётся сам COUNT. Такой совет не просто бесполезен, он сбивает с толку.
    if (-not $out.Field) {
        $out.Verdict = 'NoAggregateField'
        $out.Note    = ('{0} считает строки, а не сумму поля — SIFT для этого не нужен; помогает ключ под фильтр, а не сумма' -f $out.Function)
        return $out
    }

    # имена — как в C/AL: они пойдут в текст совета и в состав ключа
    if ($out.Field) { $out.Field = ConvertTo-KaFieldName -TableId $TableRef.TableId -Name $out.Field }
    $eq = @($Predicates.Equality | Select-Object -ExpandProperty Column -Unique | Where-Object { $_ } |
           ForEach-Object { ConvertTo-KaFieldName -TableId $TableRef.TableId -Name $_ })
    $out.NeededKeyFields = $eq

    $fLow  = ''
    if ($out.Field) { $fLow = (ConvertTo-KaSqlName $out.Field).ToLowerInvariant() }
    $eqLow = @($eq | ForEach-Object { (ConvertTo-KaSqlName $_).ToLowerInvariant() })

    $match = $null
    if ($fLow) {
        foreach ($k in (Get-KaTableKeys $TableRef.TableId)) {
            if ($k.SumLower -notcontains $fLow) { continue }
            $kf = $k.FieldsLower
            $m  = 0
            while ($m -lt $kf.Count -and ($eqLow -contains $kf[$m])) { $m++ }
            if ($m -eq $eqLow.Count -and $eqLow.Count -gt 0) { $match = $k; break }
            if (-not $match) { $match = $k }   # хотя бы какой-то ключ с этим полем суммы
        }
    }

    if ($match) {
        $pfx = ''
        if ($TableRef.Company) { $pfx = $TableRef.Company + '$' }
        $out.ExistingKeyNo = $match.KeyNo
        $out.ExistingKey   = $match
        $out.SiftViewName  = ('{0}{1}$VSIFT${2}' -f $pfx, (ConvertTo-KaSqlName $TableRef.TableName), ($match.KeyNo - 1))
        if (-not $match.Enabled) {
            $out.Verdict = 'SiftKeyDisabled'
            $out.Note    = ("SIFT-ключ {0} с суммой «{1}» есть, но выключен — представления {2} в базе нет" -f $match.KeyNo, $out.Field, $out.SiftViewName)
        }
        elseif (-not $match.MaintainSIFTIndex) {
            $out.Verdict = 'SiftNotMaintained'
            $out.Note    = ("у ключа {0} MaintainSIFTIndex=No — сумма «{1}» объявлена, а представления нет, платформа считает перебором" -f $match.KeyNo, $out.Field)
        }
        else {
            $out.Verdict = 'SiftBypassed'
            $out.Note    = ("представление {0} существует, но запрос пошёл по базовой таблице — фильтр не лёг на префикс ключа {1} ({2})" -f $out.SiftViewName, $match.KeyNo, (@($match.Fields) -join ', '))
        }
    }
    else {
        $out.Verdict = 'NoSift'
        $out.Note    = ("сумма «{0}» ни в одном SumIndexFields таблицы не объявлена — агрегат считается перебором строк" -f $out.Field)
    }
    return $out
}

# ============================================================================
#  4. New-KeyProposal — состав ключа, SumIndexFields и обязательная цена
# ============================================================================

function New-KeyProposal {
    <#
    .SYNOPSIS
        Предложение по ключу: состав, нужны ли SumIndexFields, цена и предупреждения.

    .DESCRIPTION
        Состав: сначала поля равенства, затем ОДНО поле диапазона. Поля первичного
        ключа в хвост не дописываются — SQL добавляет их в индекс сам. Больше восьми
        полей не предлагается: такой ключ дороже перебора, который он лечит. Поля из
        неиндексируемого множества сюда не попадают по построению — их отбрасывает
        Get-SqlPredicates.

        SumIndexFields предлагаются ТОЛЬКО под измеренный агрегат: без факта из замера
        это лишнее синхронное представление на каждую запись.

        Цена в ответе обязательна — она и есть содержание совета.

    .PARAMETER TableId
        Номер таблицы.

    .PARAMETER Equality
        Поля равенства, порядок сохраняется.

    .PARAMETER Range
        Поле диапазона (берётся первое).

    .PARAMETER Coverage
        Результат Test-KeyCoverage: из него берётся ключ-основа для удлинения.

    .PARAMETER AggregateField
        Поле агрегата для SumIndexFields. Учитывается только вместе с -AggregateMeasured.

    .PARAMETER AggregateMeasured
        Подтверждение, что агрегат виден в замере, а не предполагается.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int] $TableId,
        [string[]] $Equality = @(),
        [string[]] $Range    = @(),
        $Coverage,
        [string]   $AggregateField,
        [switch]   $AggregateMeasured
    )
    Initialize-KeyAdvisor

    $cls = Get-TableClass -TableId $TableId
    $tbl = ''
    if ($script:KaTables.ContainsKey($TableId)) { $tbl = $script:KaTables[$TableId] }

    $out = [pscustomobject]@{
        TableId        = $TableId
        TableName      = $tbl
        Class          = $cls.Class
        Action         = 'None'      # New | Extend | Enable | MaintainSql | None
        BaseKeyNo      = 0
        Fields         = @()
        AddedFields    = @()
        SumIndexFields = @()
        TooWide        = $false
        Skipped        = @()
        Cost           = @()
        Warnings       = @()
        Ok             = $false
    }

    # имена приводим к именам полей C/AL: их и пишут в ключ
    $Equality = @($Equality | ForEach-Object { ConvertTo-KaFieldName -TableId $TableId -Name $_ })
    $Range    = @($Range    | ForEach-Object { ConvertTo-KaFieldName -TableId $TableId -Name $_ })

    $keys  = Get-KaTableKeys $TableId
    $pkLow = @()
    $pk1   = @($keys | Where-Object { $_.KeyNo -eq 1 })
    if ($pk1.Count -gt 0) { $pkLow = $pk1[0].FieldsLower }

    # --- отбор полей, годных в индекс ---------------------------------------
    $wanted = New-Object System.Collections.Generic.List[string]
    foreach ($f in $Equality) {
        if (-not $f) { continue }
        if ($wanted -contains $f) { continue }
        $fld = Get-KaField -TableId $TableId -Name $f
        if (-not $fld) { $out.Warnings += ("поля «{0}» в справочнике таблицы нет — проверьте разбор запроса и свежесть keys/fields" -f $f) }
        if ($fld -and $fld.FieldClass -ne 'Normal') {
            $out.Skipped += ("{0} — {1}, своей колонки в SQL нет" -f $f, $fld.FieldClass); continue
        }
        if ($fld -and -not $fld.Enabled) {
            $out.Skipped += ("{0} — поле выключено" -f $f); continue
        }
        $wanted.Add($f)
    }
    if (@($Range).Count -gt 0 -and $Range[0]) {
        $fld = Get-KaField -TableId $TableId -Name $Range[0]
        if ($fld -and ($fld.FieldClass -ne 'Normal' -or -not $fld.Enabled)) {
            $out.Skipped += ("{0} — {1}, своей колонки в SQL нет" -f $Range[0], $fld.FieldClass)
        }
        elseif (-not ($wanted -contains $Range[0])) { $wanted.Add($Range[0]) }
    }

    if ($wanted.Count -eq 0) {
        $out.Warnings += 'индексируемых полей в условии не осталось — ключ не поможет, менять надо сам фильтр'
        return $out
    }

    # --- удлинить существующий или завести новый ------------------------------
    $baseKey = $null
    if ($Coverage -and $Coverage.Verdict -eq 'ExtendKey' -and $Coverage.BestKey) { $baseKey = $Coverage.BestKey }

    if ($baseKey) {
        $out.Action    = 'Extend'
        $out.BaseKeyNo = $baseKey.KeyNo
        $fields  = New-Object System.Collections.Generic.List[string]
        foreach ($f in $baseKey.Fields) { $fields.Add($f) }
        $haveLow = @($baseKey.FieldsLower)
        foreach ($f in $wanted) {
            if ($haveLow -contains (ConvertTo-KaSqlName $f).ToLowerInvariant()) { continue }
            $fields.Add($f); $out.AddedFields += $f
        }
        $out.Fields = @($fields)
    }
    elseif ($Coverage -and $Coverage.Verdict -eq 'KeyDisabled' -and $Coverage.BestKey) {
        $out.Action = 'Enable'; $out.BaseKeyNo = $Coverage.BestKey.KeyNo
        $out.Fields = @($Coverage.BestKey.Fields)
    }
    elseif ($Coverage -and $Coverage.Verdict -eq 'KeyNotMaintained' -and $Coverage.BestKey) {
        $out.Action = 'MaintainSql'; $out.BaseKeyNo = $Coverage.BestKey.KeyNo
        $out.Fields = @($Coverage.BestKey.Fields)
    }
    elseif ($Coverage -and $Coverage.Verdict -eq 'KeyExists') {
        $out.Action = 'None'
        $out.Fields = @()
    }
    else {
        $out.Action = 'New'
        # хвост из полей первичного ключа не дописываем: SQL добавит его сам
        $fields = New-Object System.Collections.Generic.List[string]
        foreach ($f in $wanted) { $fields.Add($f) }
        while ($fields.Count -gt 1 -and ($pkLow -contains (ConvertTo-KaSqlName $fields[$fields.Count - 1]).ToLowerInvariant())) {
            $out.Skipped += ("{0} — поле первичного ключа, SQL допишет его в индекс сам" -f $fields[$fields.Count - 1])
            $fields.RemoveAt($fields.Count - 1)
        }
        $out.Fields = @($fields)
    }

    if (@($out.Fields).Count -gt 8) {
        $out.TooWide   = $true
        $out.Warnings += ("в ключе получилось {0} полей — такой индекс дороже перебора, который он лечит; сузьте фильтр или разбейте выборку" -f @($out.Fields).Count)
    }

    # --- SumIndexFields только под измеренный агрегат --------------------------
    if ($AggregateField -and $AggregateMeasured) {
        $out.SumIndexFields = @($AggregateField)
    }
    elseif ($AggregateField) {
        $out.Warnings += ("агрегат по «{0}» замером не подтверждён — SumIndexFields не предлагаются" -f $AggregateField)
    }

    # --- цена: обязательная часть ответа ---------------------------------------
    if ($out.Action -ne 'None') {
        $out.Cost += 'каждая вставка, изменение и удаление строки этой таблицы будет обновлять ещё один индекс — запись дорожает ровно на него'
        $out.Cost += 'индекс занимает место и попадает в обслуживание: перестроение, статистика, резервная копия'
    }
    if (@($out.SumIndexFields).Count -gt 0) {
        $out.Cost += 'SumIndexFields создают ИНДЕКСИРОВАННОЕ ПРЕДСТАВЛЕНИЕ SIFT: оно обновляется СИНХРОННО, в той же транзакции, что и запись'
        $out.Cost += 'строка представления агрегирует много записей сразу, поэтому параллельные сессии дерутся за одни и те же строки — прямой риск блокировок и взаимоблокировок на проводке'
    }
    if ($out.Action -eq 'New' -or $out.Action -eq 'Extend' -or $out.Action -eq 'Enable' -or $out.Action -eq 'MaintainSql') {
        $out.Cost += 'смена состава ключей — синхронизация схемы: нужно окно, на большой таблице построение индекса идёт минуты и держит блокировки'
    }

    # --- обязательные предупреждения -------------------------------------------
    $out.Warnings += 'SETCURRENTKEY меняет ПОРЯДОК ОБХОДА: цикл пойдёт по строкам в другой последовательности, и бизнес-результат может стать другим. Проверять надо не скорость, а совпадение результата'
    if ($out.Action -eq 'Extend') {
        $out.Warnings += ("поля дописываются ТОЛЬКО В КОНЕЦ ключа {0}: вставка в середину переставляет порядок обхода всем, кто на этот ключ уже опирается" -f $out.BaseKeyNo)
    }
    if ($cls.Class -eq 'Standard' -and $out.Action -ne 'None') {
        $out.Warnings += 'таблица стандартная: правка уйдёт в upgrade-контур. Сначала проверить, нет ли среди ВЫКЛЮЧЕННЫХ ключей готового нужного состава — включение конфликта при обновлении не создаёт'
    }
    if ($cls.Class -eq 'ClientCore') {
        $out.Warnings += 'таблица из диапазона 10 000 000+: правка требует лицензии владельца решения. Совет передаётся владельцу подсистемы вместе с замером'
    }
    $out.Ok = (-not $out.TooWide)
    return $out
}

# ============================================================================
#  Фасад: готовый текст совета
# ============================================================================

function Format-KaFields {
    param($Fields)
    $a = @($Fields | Where-Object { $_ })
    if ($a.Count -eq 0) { return '(нет)' }
    return ($a -join ', ')
}

function Get-KaAdviceText {
    <#
    .SYNOPSIS
        Собирает формулировку совета из вердикта, предложения и проверки SIFT.
    #>
    param($Ref, $Pred, $Coverage, $Proposal, $Sift, $Class)

    $L = New-Object System.Collections.Generic.List[string]
    if ($Ref.TableId -gt 0) {
        $L.Add(("Таблица {0} «{1}» — {2}." -f $Ref.TableId, $Ref.TableName, $Class.Title))
    }
    # по таблицам, которые править нельзя, «сделай» не пишем — только сигнал
    $act = 'Действие: '
    if (-not $Class.Advise) {
        $act = 'Нужно, но не своими силами: '
        $L.Add('Политика: ' + $Class.Policy)
    }

    switch ($Coverage.Verdict) {
        'KeyExists' {
            $L.Add(("Ключ есть: {0} ({1}), индекс {2}." -f $Coverage.BestKeyNo, (Format-KaFields $Coverage.BestKey.Fields), $Coverage.SqlIndexName))
            if ($Coverage.Level -eq 'Full') { $L.Add('Фильтр и сортировка ложатся на него целиком — по ключам вопросов нет, время уходит куда-то ещё.') }
            else                            { $L.Add('Фильтр ложится на ключ, но не целиком — см. замечания ниже.') }
        }
        'KeyDisabled' {
            $L.Add(("Ключ есть, но ВЫКЛЮЧЕН: {0} ({1})." -f $Coverage.BestKeyNo, (Format-KaFields $Coverage.BestKey.Fields)))
            $L.Add((($act + "включить ключ {0} (Enabled=Yes) — состав уже нужный, новый придумывать не надо; появится индекс {1}.") -f $Coverage.BestKeyNo, $Coverage.SqlIndexName))
        }
        'KeyNotMaintained' {
            $L.Add(("Ключ подходит, но индекс НЕ ПОДДЕРЖИВАЕТСЯ: ключ {0} ({1}), MaintainSQLIndex=No." -f $Coverage.BestKeyNo, (Format-KaFields $Coverage.BestKey.Fields)))
            $L.Add($act + 'поставить MaintainSQLIndex=Yes. Сейчас ключ задаёт только порядок обхода, физического индекса под ним нет — SETCURRENTKEY на него ускорения не даст.')
        }
        'ExtendKey' {
            $L.Add(("Готового ключа нет, но ключ {0} ({1}) целиком лежит внутри фильтра — совпало полей: {2}." -f $Coverage.BestKeyNo, (Format-KaFields $Coverage.BestKey.Fields), $Coverage.MatchedPrefix))
            $L.Add((($act + "удлинить ключ {0} до «{1}» — дописать В КОНЕЦ: {2}.") -f $Coverage.BestKeyNo, (Format-KaFields $Proposal.Fields), (Format-KaFields $Proposal.AddedFields)))
        }
        'NewKey' {
            $L.Add('Подходящего ключа нет: ни один не начинается с полей этого фильтра.')
            $L.Add((($act + "новый ключ «{0}».") -f (Format-KaFields $Proposal.Fields)))
        }
        'NoPredicates' {
            $L.Add('Ни фильтра, ни сортировки — это полный перебор таблицы. Ключ не поможет: либо ограничивать выборку в коде, либо перебор тут и задуман.')
        }
        'NoIndexablePredicate' {
            $L.Add('Условия есть, но ни одно из них не ложится в индекс — читается вся таблица.')
            $L.Add($act + 'ключ не поможет, менять надо сам фильтр (см. список ниже, что именно мешает).')
        }
        'NoTable' {
            $L.Add('Ключей по этой таблице в справочнике нет — совет не даётся.')
        }
    }

    if ($Coverage.Verdict -eq 'ExtendKey' -or $Coverage.Verdict -eq 'NewKey') {
        $L.Add(("Из запроса: равенство — {0}; диапазон — {1}; сортировка — {2}." -f
            (Format-KaFields $Coverage.Equality), (Format-KaFields $Coverage.Range), (Format-KaFields $Coverage.OrderBy)))
    }
    if (@($Proposal.SumIndexFields).Count -gt 0) {
        $L.Add(("SumIndexFields: {0} — под измеренный агрегат." -f (Format-KaFields $Proposal.SumIndexFields)))
    }
    if ($Sift -and $Sift.IsAggregate -and $Sift.Verdict -ne 'NotAggregate') {
        $L.Add(("SIFT: {0}." -f $Sift.Note))
    }

    foreach ($n in @($Coverage.Notes))   { $L.Add(("Замечание: {0}." -f $n)) }
    foreach ($s in @($Proposal.Skipped)) { $L.Add(("Не берём в ключ: {0}." -f $s)) }
    foreach ($n in @($Pred.NonIndexable)) {
        $c = $n.Column; if (-not $c) { $c = '(группа условий)' }
        $L.Add(("В ключ не идёт: {0} {1} — {2}." -f $c, $n.Op, $n.Reason))
    }
    foreach ($u in @($Pred.Unknown)) {
        $c = $u.Column; if (-not $c) { $c = '(условие)' }
        $L.Add(("Разбор ненадёжен: {0} {1} — {2}." -f $c, $u.Op, $u.Reason))
    }

    if (@($Proposal.Cost).Count -gt 0) {
        $L.Add('Цена:')
        foreach ($c in @($Proposal.Cost)) { $L.Add(("  - {0}." -f $c)) }
    }
    if (@($Proposal.Warnings).Count -gt 0) {
        $L.Add('Внимание:')
        foreach ($w in @($Proposal.Warnings)) { $L.Add(("  - {0}." -f $w)) }
    }
    return ($L -join "`r`n")
}

function Invoke-KeyAdvisor {
    <#
    .SYNOPSIS
        Полный проход по одному запросу: таблица -> предикаты -> покрытие -> совет.

    .DESCRIPTION
        Возвращает вердикт и готовый текст совета. Результат кешируется по тексту
        запроса: в трассировке один и тот же оператор повторяется тысячами раз.

    .PARAMETER Sql
        Текст запроса как пришёл из события; хвост AppObjectType/AppObjectId снимать
        не нужно.

    .PARAMETER Measured
        Признак «агрегат подтверждён замером»: только с ним предлагаются SumIndexFields.

    .PARAMETER NoCache
        Не брать и не класть в кеш.

    .OUTPUTS
        Verdict: KeyExists | KeyDisabled | KeyNotMaintained | ExtendKey | NewKey |
                 NoIndexablePredicate | NoPredicates | NoTable | Ambiguous |
                 Unresolved | NotAQuery | Silent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Sql,
        [switch] $Measured,
        [switch] $NoCache
    )
    Initialize-KeyAdvisor
    $ck = $Sql
    if ($Measured) { $ck = 'M|' + $ck }
    if (-not $NoCache -and $script:KaCache.ContainsKey($ck)) { return $script:KaCache[$ck] }

    $ref  = Get-SqlTableRef -Sql $Sql
    $pred = Get-SqlPredicates -Sql $Sql -Qualifier $ref.Alias
    $cls  = Get-TableClass -TableId $ref.TableId -SqlName $ref.SqlName

    $res = [pscustomobject]@{
        Ok         = $false
        Verdict    = 'Unresolved'
        Table      = $ref
        Predicates = $pred
        Class      = $cls
        Coverage   = $null
        Proposal   = $null
        Sift       = $null
        Advice     = ''
    }

    # порядок важен: служебная таблица — молчим; не разобрали — так и говорим,
    # а не прячем под тем же «молчим»
    if ($ref.Status -eq 'NotAQuery') {
        $res.Verdict = 'NotAQuery'
        $res.Ok      = $true
        if (-not $NoCache) { $script:KaCache[$ck] = $res }
        return $res
    }
    if ($ref.Status -eq 'System') {
        $res.Verdict = 'Silent'
        $res.Ok      = $true
        if (-not $NoCache) { $script:KaCache[$ck] = $res }
        return $res
    }
    if (-not $ref.Ok) {
        if ($ref.Status -eq 'Ambiguous') { $res.Verdict = 'Ambiguous' } else { $res.Verdict = 'Unresolved' }
        $res.Advice = ('Таблица по запросу не определена: ' + $ref.Reason + '. Совет не даётся.')
        if (-not $NoCache) { $script:KaCache[$ck] = $res }
        return $res
    }
    if ($cls.Silent) {
        $res.Verdict = 'Silent'
        $res.Ok      = $true
        if (-not $NoCache) { $script:KaCache[$ck] = $res }
        return $res
    }

    $eq = @($pred.Equality | Select-Object -ExpandProperty Column -Unique | Where-Object { $_ })
    $rg = @($pred.Range    | Select-Object -ExpandProperty Column -Unique | Where-Object { $_ })

    $other = ((@($pred.NonIndexable).Count + @($pred.Unknown).Count) -gt 0)
    $cov   = Test-KeyCoverage -TableId $ref.TableId -Equality $eq -Range $rg -OrderBy $pred.OrderBy `
                             -Company $ref.Company -HasOtherPredicates:$other
    $sift = Test-SiftCoverage -TableRef $ref -Predicates $pred

    $aggField = ''
    if ($sift.IsAggregate -and $sift.Verdict -eq 'NoSift') { $aggField = $sift.Field }

    $prop = New-KeyProposal -TableId $ref.TableId -Equality $eq -Range $rg -Coverage $cov `
                            -AggregateField $aggField -AggregateMeasured:$Measured

    $res.Coverage = $cov
    $res.Proposal = $prop
    $res.Sift     = $sift
    $res.Verdict  = $cov.Verdict
    $res.Advice   = Get-KaAdviceText -Ref $ref -Pred $pred -Coverage $cov -Proposal $prop -Sift $sift -Class $cls
    $res.Ok       = $true

    if (-not $NoCache) { $script:KaCache[$ck] = $res }
    return $res
}
