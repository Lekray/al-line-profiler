<#
.SYNOPSIS
    Движок подсказок по оптимизации C/AL: каталог правил поверх статического разбора
    (Lib-AlParse) с гейтом по измеренной цене строки (метрики Lib-AlTrace).

.DESCRIPTION
    Главный принцип: правило = синтаксический паттерн + гейт по цене. Линтер по
    1,35 млн строк без гейта выдал бы тысячи находок на холодном коде. Поэтому:
      - с метриками (-MetricsFile) находка получает уровень по замеру строки:
        High — строка дорогая (>= -HotLineMs мс либо >= -HotLinePct % объекта
        и не меньше -MinHotMs),
        Medium — строка исполнялась, но дешёвая; Info — холодная или не исполнялась;
      - без метрик ВСЕ находки понижаются до Info (статический прогон);
      - мёртвый код тонет сам: вклад ноль — приоритет ноль.

    Каталог правил.
    Построчные:
      101 COUNT для проверки пустоты вместо ISEMPTY
      102 легаси FIND('-') / FIND('+')
      103 SETFILTER с отрицанием или маской слева — не ложится в индекс
      104 фильтр по FlowField (точный факт метаданных, fields.tsv)
      110 CALCFIELDS в цикле
      111 GET в цикле (N+1)
      112 запрос внутри цикла (N+1); инвариантный к итерации — вынести из цикла
      113 построчные MODIFY/INSERT/DELETE в цикле
      114 COMMIT в цикле
      115 внешний вызов в цикле (Automation/OCX, тяжёлый DotNet, SLEEP/SHELL)
      116 повторный одинаковый запрос без смены фильтров
      118 повторное чтение BLOB (CALCFIELDS поля типа BLOB)
      119 LOCKTABLE после чтения
    Уровня области (только с метриками):
      201 горячая функция
      202 кардинальность цикла (итерации по Hits)
      203 диагноз: функция упирается в SQL или в C/AL
      204 дважды вычисленное одинаковое выражение рядом (работает и статически)
    Ключи (делегируется Lib-KeyAdvisor):
      301 по цепочке SETRANGE/SETFILTER/SETCURRENTKEY -> потребитель строится
          состав фильтра и сверяется с ключами таблицы: NewKey / ExtendKey /
          KeyDisabled / KeyNotMaintained; для CALCSUMS — проверка SumIndexFields.

    Сквозные правила:
      - TEMPORARY-переменные пропускаются всеми SQL-правилами; двойная проверка:
        флаг из таблицы символов И отсутствие SQL-событий на строке в замере;
      - подавление ложных срабатываний — по СИГНАТУРЕ (правило + функция +
        нормализованный текст строки), а не по номеру строки: файл -SuppressFile.

    Формат вывода — семь колонок статического анализатора C/AL плюс Severity:
        Строка  RuleID  Тип  ID  Функция  Элемент  Сообщение  Severity
    Этот же файл читает Build-Report.ps1 параметром -HintsFile.

.PARAMETER ObjectType
    Тип объекта (1 Table, 3 Report, 5 Codeunit, 6 XMLport, 8 Page, 9 Query).

.PARAMETER ObjectId
    Номер объекта.

.PARAMETER All
    Пройти по всем объектам дампа .alsrc (с фильтрами -Types и -Top).

.PARAMETER Top
    С -All: взять N крупнейших объектов по числу строк. 0 — все.

.PARAMETER Types
    С -All: только эти типы объектов.

.PARAMETER MetricsFile
    lines.tsv от Lib-AlTrace (Export-AlLineMetrics). Необязателен: без него —
    статический прогон, все находки Info.

.PARAMETER MinSeverity
    Info | Medium | High. По умолчанию: Info без метрик, Medium с метриками
    (гейт по цене отсекает холодные строки).

.PARAMETER OutFile
    TSV с находками. По умолчанию out\hints.tsv (один объект) или
    out\perflint-all.tsv (-All).

.PARAMETER SuppressFile
    TSV подавлений: RuleID, Тип, ID, Функция, Сигнатура (верхний регистр
    нормализованного текста строки). По умолчанию out\perflint-suppress.tsv,
    если существует.

.PARAMETER HotLineMs
    Порог «строка дорогая», мс собственного времени за прогон. По умолчанию 10.

.PARAMETER HotLinePct
    Порог «строка дорогая», % собственного времени объекта. По умолчанию 1.

.PARAMETER HotFuncMs
    Порог «горячая функция», мс. По умолчанию 50.

.PARAMETER HotFuncPct
    Порог «горячая функция», % собственного времени объекта. По умолчанию 10.

.PARAMETER MinHotMs
    Нижний абсолютный порог, мс: без него процентный гейт на коротком прогоне
    поднимал бы в High строки в доли миллисекунды — это шум на уровне разрешения
    таймера. По умолчанию 1.

.PARAMETER LoopHits
    Порог итераций цикла для правила 202, он же — порог, начиная с которого
    строка считается лежащей в цикле по одному лишь замеру (когда цикла в тексте
    нет: перебор ведёт платформа). По умолчанию 100.

.EXAMPLE
    .\Invoke-PerfLint.ps1 -ObjectType 5 -ObjectId 80
    Статический прогон по Sales-Post: все находки Info.

.EXAMPLE
    .\Invoke-PerfLint.ps1 -ObjectType 5 -ObjectId 80 -MetricsFile ..\out\lines.tsv
    С метриками: показываются только строки, дорогие в замере.

.EXAMPLE
    .\Invoke-PerfLint.ps1 -All -Top 250 -OutFile ..\out\perflint-all.tsv
    Массовый прогон по 250 крупнейшим объектам.
#>
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'One')]
param(
    [Parameter(ParameterSetName = 'One', Mandatory)][int] $ObjectType,
    [Parameter(ParameterSetName = 'One', Mandatory)][int] $ObjectId,
    [Parameter(ParameterSetName = 'All', Mandatory)][switch] $All,
    [Parameter(ParameterSetName = 'All')][int] $Top = 0,
    [Parameter(ParameterSetName = 'All')][int[]] $Types,
    [string] $MetricsFile,
    [ValidateSet('Info', 'Medium', 'High')][string] $MinSeverity,
    [string] $OutFile,
    [string] $SuppressFile,
    [string] $SourceRoot,
    [string] $BaseRoot,
    [double] $HotLineMs  = 10.0,
    [double] $HotLinePct = 1.0,
    [double] $HotFuncMs  = 50.0,
    [double] $HotFuncPct = 10.0,
    [double] $MinHotMs   = 1.0,
    [int]    $LoopHits   = 100,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Lib-AlParse.ps1')      # тянет и Lib-AlListing.ps1
. (Join-Path $PSScriptRoot 'Lib-KeyAdvisor.ps1')

$script:PlOutDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'out'

# ---------------------------------------------------------------------------
# Метрики: lines.tsv от Export-AlLineMetrics, колонки по именам
# ---------------------------------------------------------------------------

function Read-PlMetrics {
    param([string] $Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Нет файла метрик: $Path" }
    $inv   = [System.Globalization.CultureInfo]::InvariantCulture
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    if ($lines.Length -lt 2) { throw "Файл метрик пуст: $Path" }
    $head = $lines[0] -split "`t"
    $ix = @{}
    for ($i = 0; $i -lt $head.Length; $i++) { $ix[$head[$i].Trim()] = $i }
    foreach ($req in @('ObjectType', 'ObjectId', 'LineNo', 'SelfMs')) {
        if (-not $ix.ContainsKey($req)) { throw "В метриках нет колонки $req" }
    }
    $objs = @{}
    for ($i = 1; $i -lt $lines.Length; $i++) {
        if (-not $lines[$i]) { continue }
        $c = $lines[$i] -split "`t"
        $getS = { param($n) if ($ix.ContainsKey($n) -and $ix[$n] -lt $c.Length) { $c[$ix[$n]] } else { '' } }
        $getD = { param($n) $v = & $getS $n; $d = 0.0
                  if ($v) { [void][double]::TryParse($v, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d) }; $d }
        $getI = { param($n) $v = & $getS $n; $r = 0L; if ($v) { [void][int64]::TryParse($v, [ref]$r) }; $r }
        $key = ('{0}/{1}' -f (& $getS 'ObjectType'), (& $getS 'ObjectId'))
        if (-not $objs.ContainsKey($key)) {
            $objs[$key] = @{
                Lines = @{}; FuncSelf = @{}; FuncSql = @{}; FuncHits = @{}
                TotalSelf = 0.0; TotalSql = 0.0
            }
        }
        $o  = $objs[$key]
        $ln = [int](& $getI 'LineNo')
        $fn = & $getS 'FunctionName'
        $row = [pscustomobject]@{
            LineNo   = $ln
            Function = $fn
            Hits     = (& $getI 'Hits')
            TotalMs  = (& $getD 'TotalMs')
            SelfMs   = (& $getD 'SelfMs')
            SqlMs    = (& $getD 'SqlMs')
            SqlCount = (& $getI 'SqlCount')
        }
        $o.Lines[$ln]  = $row
        $o.TotalSelf  += $row.SelfMs
        $o.TotalSql   += $row.SqlMs
        if ($fn) {
            if (-not $o.FuncSelf.ContainsKey($fn)) { $o.FuncSelf[$fn] = 0.0; $o.FuncSql[$fn] = 0.0; $o.FuncHits[$fn] = 0L }
            $o.FuncSelf[$fn] += $row.SelfMs
            $o.FuncSql[$fn]  += $row.SqlMs
            if ($row.Hits -gt $o.FuncHits[$fn]) { $o.FuncHits[$fn] = $row.Hits }
        }
    }
    return $objs
}

# ---------------------------------------------------------------------------
# Гейт по цене: уровень находки из замера строки
# ---------------------------------------------------------------------------

function Get-PlHeat {
    <#
    .SYNOPSIS
        Уровень находки по метрикам строки: High / Medium / Info + оценка выигрыша.
    #>
    param($M, [int] $Line, [string] $Func)
    if ($null -eq $M) {
        return [pscustomobject]@{ Sev = 'Info'; GainMs = -1.0; Note = 'без замера' }
    }
    if ($M.Lines.ContainsKey($Line)) {
        $r = $M.Lines[$Line]
        $v = $r.SelfMs
        $hot = ($v -ge $HotLineMs)
        # Доля от прогона поднимает строку в High только вместе с абсолютным полом:
        # на коротком замере 1 % — это доли миллисекунды, то есть шум на уровне
        # разрешения таймера, и такая строка не стоит внимания разработчика.
        if (-not $hot -and $M.TotalSelf -gt 0 -and
            $v -ge ($M.TotalSelf * $HotLinePct / 100.0) -and $v -ge $MinHotMs) { $hot = $true }
        if ($hot) {
            return [pscustomobject]@{ Sev = 'High'; GainMs = $v
                Note = ('в замере: {0:0.#} мс, вызовов {1}, SQL {2}' -f $v, $r.Hits, $r.SqlCount) }
        }
        if ($v -gt 0 -or $r.Hits -gt 0) {
            return [pscustomobject]@{ Sev = 'Medium'; GainMs = $v
                Note = ('в замере: {0:0.##} мс, вызовов {1}' -f $v, $r.Hits) }
        }
        return [pscustomobject]@{ Sev = 'Info'; GainMs = 0.0; Note = 'строка в замере холодная' }
    }
    # своего замера у строки нет: многострочный оператор или строка не исполнялась
    if ($Func -and $M.FuncSelf.ContainsKey($Func)) {
        if ($M.FuncSelf[$Func] -ge $HotFuncMs) {
            return [pscustomobject]@{ Sev = 'Medium'; GainMs = 0.0
                Note = ('своего замера строки нет, функция горячая: {0:0.#} мс' -f $M.FuncSelf[$Func]) }
        }
        return [pscustomobject]@{ Sev = 'Info'; GainMs = 0.0; Note = 'строка в замере не исполнялась' }
    }
    return [pscustomobject]@{ Sev = 'Info'; GainMs = 0.0; Note = 'функция в замер не попала' }
}

# ---------------------------------------------------------------------------
# Мелкие помощники
# ---------------------------------------------------------------------------

function Remove-PlLineComment {
    <# .SYNOPSIS Срезает //-комментарий вне строкового литерала (литералы сохраняются). #>
    param([string] $Text)
    if (-not $Text -or $Text.IndexOf('/') -lt 0) { return $Text }
    $inStr = $false
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($inStr) { if ($ch -eq "'") { $inStr = $false }; continue }
        if ($ch -eq "'") { $inStr = $true; continue }
        if ($ch -eq '/' -and $i + 1 -lt $Text.Length -and $Text[$i + 1] -eq '/') {
            return $Text.Substring(0, $i)
        }
    }
    return $Text
}

function Split-PlArgs {
    <#
    .SYNOPSIS
        Делит список аргументов по запятым верхнего уровня (кавычки и скобки учитываются).

    .DESCRIPTION
        КОНТРАКТ: результат возвращается одним объектом («return ,$arr»). Присваивайте
        его в переменную — при присваивании обёртка снимается и получается массив.
        НЕ подавайте вызов прямо в конвейер: конвейер снимет ровно обёртку и передаст
        дальше весь массив ОДНИМ элементом, который потом молча склеится в строку
        через пробел. Симптом — «Debit Amount" "Credit Amount» вместо двух полей.
    #>
    param([string] $Text)
    $parts = New-Object System.Collections.ArrayList
    if ($null -eq $Text) { return ,$parts.ToArray() }
    $depth = 0; $inSq = $false; $inDq = $false; $start = 0
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($inSq) { if ($ch -eq "'") { $inSq = $false }; continue }
        if ($inDq) { if ($ch -eq '"') { $inDq = $false }; continue }
        switch ($ch) {
            "'" { $inSq = $true }
            '"' { $inDq = $true }
            '(' { $depth++ }
            '[' { $depth++ }
            ')' { if ($depth -gt 0) { $depth-- } }
            ']' { if ($depth -gt 0) { $depth-- } }
            ',' { if ($depth -eq 0) { [void]$parts.Add($Text.Substring($start, $i - $start).Trim()); $start = $i + 1 } }
        }
    }
    if ($start -le $Text.Length) { [void]$parts.Add($Text.Substring($start).Trim()) }
    return ,$parts.ToArray()
}

function Get-PlCallArgText {
    <#
    .SYNOPSIS
        Текст аргументов вызова Var.Op(...) из СЫРОГО листинга (литералы целы).
    .DESCRIPTION
        Вызов может занимать несколько строк — дочитываем до баланса скобок
        (до 12 строк), с каждой строки предварительно срезав //-комментарий.
        -Occur выбирает N-е вхождение оператора на строке (нумерация с 0).
    #>
    param($Rows, [int] $LineNo, [string] $Var, [string] $Op, [int] $Occur = 0)
    $n = $Rows.Count
    if ($LineNo -lt 1 -or $LineNo -gt $n) { return $null }
    $sb = New-Object System.Text.StringBuilder
    $last = [math]::Min($n, $LineNo + 12)
    for ($l = $LineNo; $l -le $last; $l++) {
        [void]$sb.Append((Remove-PlLineComment $Rows[$l - 1].Text))
        [void]$sb.Append(' ')
    }
    $joined = $sb.ToString()
    if ($Var) {
        $ve = [regex]::Escape($Var)
        $pat = ('(?:"{0}"|(?<![\w"]){0})\s*\.\s*{1}\s*\(' -f $ve, $Op)
    }
    else {
        $pat = ('(?<![\w".]){0}\s*\(' -f $Op)
    }
    $ms = [regex]::Matches($joined, $pat, 'IgnoreCase')
    if ($ms.Count -le $Occur) { return $null }
    $m = $ms[$Occur]
    $openIx = $m.Index + $m.Length - 1
    return (Get-AlParenSpan ($joined.Substring($openIx)))
}

function Get-PlLiteral {
    <# .SYNOPSIS Если аргумент — строковый литерал 'x', возвращает x; иначе $null. #>
    param([string] $Arg)
    if ($null -eq $Arg) { return $null }
    $m = [regex]::Match($Arg.Trim(), "^'(?<v>(?:[^']|'')*)'$")
    if ($m.Success) { return $m.Groups['v'].Value.Replace("''", "'") }
    return $null
}

function Get-PlFieldArg {
    <# .SYNOPSIS Имя поля из аргумента SETRANGE/SETFILTER/CALCFIELDS: снимаются кавычки. #>
    param([string] $Arg)
    if ($null -eq $Arg) { return '' }
    $s = $Arg.Trim()
    if ($s.Length -ge 2 -and $s[0] -eq '"' -and $s[$s.Length - 1] -eq '"') { return $s.Substring(1, $s.Length - 2) }
    return $s
}


# ---------------------------------------------------------------------------
# Символы и разрешение переменных
# ---------------------------------------------------------------------------

function New-PlSymIndex {
    param($Symbols)
    $idx = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($s in $Symbols) { $idx[($s.Scope + '|' + $s.Name)] = $s }
    return $idx
}

function Resolve-PlVar {
    <#
    .SYNOPSIS
        Переменная -> {Known, IsRecord, TableNo, TableName, Temporary}.
    .DESCRIPTION
        Порядок поиска: локальная область функции, затем глобальная. Rec/xRec
        в таблице разрешаются в саму таблицу. Пустой индекс символов (нет экспорта
        в baseline) даёт Known=false — правила обязаны понижать уверенность или молчать.
    #>
    param([string] $Var, [string] $Func, $SymIdx, [int] $ObjType, [int] $ObjId)
    $r = [pscustomobject]@{ Known = $false; IsRecord = $false; TableNo = 0; TableName = ''; Temporary = $false }
    if (-not $Var) { return $r }
    if (($Var -eq 'Rec' -or $Var -eq 'xRec') -and $ObjType -eq 1 -and $ObjId -gt 0) {
        $r.Known = $true; $r.IsRecord = $true; $r.TableNo = $ObjId
        $r.TableName = Get-AlTableNameById -TableId $ObjId
        return $r
    }
    if ($null -eq $SymIdx -or $SymIdx.Count -eq 0) { return $r }
    $sym = $null
    if (-not $SymIdx.TryGetValue(($Func + '|' + $Var), [ref]$sym)) {
        [void]$SymIdx.TryGetValue(('|' + $Var), [ref]$sym)
    }
    if ($null -eq $sym) { return $r }
    $r.Known = $true
    if ($sym.Kind -eq 'Record') {
        $r.IsRecord  = $true
        if ($null -ne $sym.ObjectId) { $r.TableNo = [int]$sym.ObjectId }
        $r.TableName = $sym.TableName
        $r.Temporary = [bool]$sym.Temporary
    }
    return $r
}

function Test-PlTempSkip {
    <#
    .SYNOPSIS
        Сквозное правило TEMPORARY: пропустить SQL-правило, если переменная временная
        И замер строки не показывает SQL (двойная проверка).
    #>
    param($VarInfo, $M, [int] $Line)
    if ($null -eq $VarInfo -or -not $VarInfo.Temporary) { return $false }
    if ($null -ne $M -and $M.Lines.ContainsKey($Line) -and $M.Lines[$Line].SqlCount -gt 0) {
        return $false   # символ говорит TEMPORARY, а SQL в замере есть — baseline отстал, не молчим
    }
    return $true
}

# ---------------------------------------------------------------------------
# Находки: подавление по сигнатуре, сборка сообщения
# ---------------------------------------------------------------------------

$script:PlSuppress   = $null   # HashSet сигнатур
$script:PlSuppressed = 0

function Read-PlSuppress {
    param([string] $Path)
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return ,$set }
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $ln = $lines[$i].Trim()
        if (-not $ln -or $ln.StartsWith('#')) { continue }
        $c = $ln -split "`t"
        if ($c.Length -lt 5) { continue }
        if ($i -eq 0 -and $c[0] -match '^RuleID$') { continue }   # заголовок
        [void]$set.Add(('{0}|{1}|{2}|{3}|{4}' -f $c[0].Trim(), $c[1].Trim(), $c[2].Trim(),
            $c[3].Trim(), $c[4].Trim().ToUpperInvariant()))
    }
    return ,$set
}

function Get-PlSignature {
    <# .SYNOPSIS Сигнатура строки для подавления: ВЕРХНИЙ(очищенный текст). #>
    param($Lex, [int] $Line)
    if ($Line -ge 1 -and $Line -le $Lex.Count) { return $Lex[$Line - 1].Clean.ToUpperInvariant() }
    return ''
}

function New-PlContext {
    param([int] $ObjType, [int] $ObjId, [string] $TypeName, $Lex, $M)
    return @{
        ObjType = $ObjType; ObjId = $ObjId; TypeName = $TypeName
        Lex = $Lex; M = $M
        Findings = New-Object System.Collections.ArrayList
    }
}

function Add-PlFinding {
    <#
    .SYNOPSIS
        Регистрирует находку: гейт по цене, подавление по сигнатуре, сборка сообщения.
    #>
    param(
        $Ctx,
        [Parameter(Mandatory)][int]    $Line,
        [Parameter(Mandatory)][string] $Rule,
        [string] $Func,
        [string] $Elem,
        [Parameter(Mandatory)][string] $Msg,
        [ValidateSet('высокая','средняя','низкая','точная')][string] $Conf = 'средняя',
        [string] $Severity,          # явный уровень (правила 201-203); иначе — гейт
        [double] $GainMs = -1.0
    )
    $heat = $null
    if (-not $Severity) {
        $heat = Get-PlHeat -M $Ctx.M -Line $Line -Func $Func
        $Severity = $heat.Sev
        if ($GainMs -lt 0) { $GainMs = $heat.GainMs }
    }

    # подавление по сигнатуре (функция + нормализованный текст строки)
    $sig = ('{0}|{1}|{2}|{3}|{4}' -f $Rule, $Ctx.ObjType, $Ctx.ObjId, $Func, (Get-PlSignature $Ctx.Lex $Line))
    if ($script:PlSuppress -and $script:PlSuppress.Contains($sig)) { $script:PlSuppressed++; return }

    $gainTxt = 'выигрыш не измерен'
    if ($GainMs -ge 0) {
        if ($GainMs -gt 0) { $gainTxt = ('выигрыш: до {0:0.#} мс за прогон' -f $GainMs) }
        else               { $gainTxt = 'вклад в замере: 0' }
    }
    $note = ''
    if ($heat -and $heat.Note -and $heat.Note -ne 'без замера') { $note = '; ' + $heat.Note }
    $full = ('{0} [уверенность: {1}; {2}{3}]' -f $Msg, $Conf, $gainTxt, $note)
    $full = $full.Replace("`r", ' ').Replace("`n", ' ').Replace("`t", ' ')

    if (-not $Elem) {
        if ($Line -ge 1 -and $Line -le $Ctx.Lex.Count) { $Elem = $Ctx.Lex[$Line - 1].Clean }
    }
    if ($Elem.Length -gt 40) { $Elem = $Elem.Substring(0, 38) + '..' }
    $Elem = $Elem.Replace("`t", ' ')

    [void]$Ctx.Findings.Add([pscustomobject]@{
        LineNo   = $Line
        RuleId   = $Rule
        TypeName = $Ctx.TypeName
        ObjectId = $Ctx.ObjId
        Function = $Func
        Element  = $Elem
        Message  = $full
        Severity = $Severity
        GainMs   = $GainMs
    })
}

# ---------------------------------------------------------------------------
# Справочники правил
# ---------------------------------------------------------------------------

# встроенные функции C/AL и методы Record: правило 204 их не трогает
$script:PlBuiltins = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($b in @(
    'ABS','ARRAYLEN','ASCII','CALCDATE','CLEAR','CLEARALL','CLOSINGDATE','CONFIRM','COPY',
    'COPYARRAY','COPYSTR','CONVERTSTR','CREATE','CREATEDATETIME','CREATEGUID','CURRENTDATETIME',
    'DATE2DMY','DATE2DWY','DELCHR','DELSTR','DMY2DATE','DWY2DATE','DT2DATE','DT2TIME','ERROR',
    'EVALUATE','EXIT','FORMAT','GETLASTERRORTEXT','INCSTR','INSSTR','LOWERCASE','MAXSTRLEN',
    'MESSAGE','NORMALDATE','PADSTR','POWER','RANDOM','RANDOMIZE','ROUND','SELECTSTR','SLEEP',
    'STRCHECKSUM','STRLEN','STRMENU','STRPOS','STRSUBSTNO','TIME','TODAY','UPPERCASE','WORKDATE',
    'SETRANGE','SETFILTER','SETCURRENTKEY','GET','FIND','FINDSET','FINDFIRST','FINDLAST','NEXT',
    'INSERT','MODIFY','DELETE','MODIFYALL','DELETEALL','RENAME','VALIDATE','TESTFIELD',
    'FIELDERROR','INIT','RESET','CALCFIELDS','CALCSUMS','SETAUTOCALCFIELDS','COPYFILTER',
    'COPYFILTERS','GETFILTER','GETFILTERS','GETRANGEMIN','GETRANGEMAX','SETRECFILTER',
    'TRANSFERFIELDS','LOCKTABLE','CHANGECOMPANY','SETPOSITION','GETPOSITION','FIELDCAPTION',
    'TABLECAPTION','FIELDNO','COUNT','COUNTAPPROX','ISEMPTY','MARK','MARKEDONLY','CLEARMARKS',
    'RUN','RUNMODAL','SETTABLEVIEW','SETRECORD','GETRECORD','LOOKUPMODE','UPDATE','CLOSE',
    'OPEN','READ','WRITE','SEEK','LEN','POS','QUERYCLOSE','SETSELECTIONFILTER','COPYLINKS',
    'HYPERLINK','DOWNLOAD','UPLOAD','SHELL','CODEUNIT','PAGE','REPORT','XMLPORT','QUERY',
    # ключевые слова C/AL: скобка после них — группировка выражения, не вызов
    'IF','THEN','ELSE','CASE','OF','WHILE','UNTIL','REPEAT','FOR','TO','DOWNTO','DO',
    'WITH','BEGIN','END','AND','OR','NOT','XOR','DIV','MOD','IN','TRUE','FALSE',
    # частые методы DotNet-обёрток: дёшевы, повтор — не находка
    'TOSTRING','APPEND','APPENDLINE','ADD','ITEM','CONTAINS','SPLIT','SUBSTRING','TRIM',
    'INDEXOF','PARSE','GETTYPE','SELECTLATESTVERSION','COMMIT'
)) { [void]$script:PlBuiltins.Add($b) }

# тяжёлые DotNet-типы для правила 115 (сетевые, процессные, межпроцессные обёртки)
$script:PlHeavyDotNetRx = [regex]'(?i)(Http|WebRequest|WebClient|WebService|Smtp|Ftp|Socket|Sql|Odbc|OleDb|Process|ServiceController|Shell|XMLHTTP|WinHttp)'

# единый детектор операторов на очищенной строке
$script:PlLineOpRx = [regex]('(?<![\w".])(?:(?<var>"[^"]+"|[A-Za-z_][A-Za-z0-9_]*)\s*\.\s*)?' +
    '(?<op>CALCFIELDS|CALCSUMS|MODIFYALL|DELETEALL|MODIFY|INSERT|DELETE|GET|FINDSET|FINDFIRST|FINDLAST|FIND|LOCKTABLE|COMMIT|SLEEP|SHELL|COUNT)(?![\w])')

# COUNT в сравнении с нулём/единицей
$script:PlCountCmpRx = [regex]('(?<![\w".])(?:(?<var>"[^"]+"|[A-Za-z_][A-Za-z0-9_]*)\s*\.\s*)?' +
    'COUNT\s*(?<op>=|<>|<=|>=|<|>)\s*(?<n>[01])(?![\w.])')

# вызов-выражение для правила 204 (скобки до двух уровней вложенности)
$script:PlCallRx = [regex]('(?<head>(?:(?:"[^"]+"|[A-Za-z_][A-Za-z0-9_]*)\s*\.\s*)?(?:"[^"]+"|[A-Za-z_][A-Za-z0-9_]*))' +
    '\s*\((?<args>[^()]*(?:\([^()]*(?:\([^()]*\)[^()]*)*\)[^()]*)*)\)')


# ---------------------------------------------------------------------------
# Движок по одному объекту
# ---------------------------------------------------------------------------

function Get-PlImplicitRecord {
    <#
    .SYNOPSIS
        Запись, к которой относится вызов без имени переменной.

    .DESCRIPTION
        В таблице это Rec. В отчёте и XMLport'е безымянный вызов внутри триггера
        DataItem'а адресует запись этого DataItem'а, а имя DataItem'а стоит в имени
        триггера: «G/L Account - OnAfterGetRecord». Без этого сообщение находки
        выходило с дырой на месте записи («CALCFIELDS(...) по  в цикле»).

        В XMLport'е к имени триггера добавлено направление: не «X - OnAfterGetRecord»,
        а «X - Export::OnAfterGetRecord». Без учёта этого разбор молча не срабатывал.

        На странице безымянный вызов идёт к Rec. Таблицу подтвердить нечем, поэтому
        правила, которым нужен номер таблицы, дальше сами промолчат, — но само имя
        записи в тексте находки должно стоять, иначе выходит «CALCFIELDS(...) по  в цикле».
    #>
    param([int]$ObjType, [string]$Func)
    if ($ObjType -eq 1 -or $ObjType -eq 8) { return 'Rec' }
    if (($ObjType -eq 3 -or $ObjType -eq 6) -and $Func) {
        $m = [regex]::Match($Func, '^(?<di>.+?)\s-\s(?:Export::|Import::)?(?:OnPreDataItem|OnAfterGetRecord|OnPostDataItem)$')
        if ($m.Success) { return $m.Groups['di'].Value }
    }
    return ''
}

function Test-PlImplicitLoop {
    <#
    .SYNOPSIS
        Триггер, который платформа сама вызывает на каждой записи набора.

    .DESCRIPTION
        В отчётах и XMLport'ах перебор DataItem'а ведёт платформа: «<DataItem> -
        OnAfterGetRecord» отрабатывает на каждой записи, а синтаксического цикла в
        коде нет. То же у OnAfterGetRecord страницы — он срабатывает на каждой
        показанной строке. Для правил «... в цикле» такой триггер равнозначен телу
        REPEAT: находка внутри умножается на число записей.
    #>
    param([int]$ObjType, [string]$Func)
    if ($ObjType -ne 3 -and $ObjType -ne 6 -and $ObjType -ne 8) { return $false }
    if (-not $Func) { return $false }
    # «X - Export::OnAfterGetRecord» у XMLport'а, «X - OnAfterGetRecord» у отчёта,
    # просто «OnAfterGetRecord» у страницы.
    return ($Func -match '(?:^|\s-\s)(?:Export::|Import::)?OnAfterGetRecord$')
}

function Invoke-PlObject {
    param(
        [int] $ObjType, [int] $ObjId,
        $MetricsAll,
        [string] $SrcRoot, [string] $BaseDir
    )
    $rows = ConvertTo-AlRows (Get-AlListing -ObjectType $ObjType -ObjectId $ObjId -SourceRoot $SrcRoot)
    $lex  = ConvertTo-AlRows (Get-AlLexed -Listing $rows)
    $struct = Get-AlStructure -Listing $rows -Lexed $lex
    $syms   = ConvertTo-AlRows (Get-AlSymbols -ObjectType $ObjType -ObjectId $ObjId `
                -BaseRoot $BaseDir -SourceRoot $SrcRoot -WarningAction SilentlyContinue)
    $chains = ConvertTo-AlRows (Get-AlFilterChains -Listing $rows -Lexed $lex -Structure $struct `
                -Symbols $syms -ObjectType $ObjType -ObjectId $ObjId)
    $symIdx = New-PlSymIndex $syms

    $met = $null
    if ($MetricsAll) {
        $k = '{0}/{1}' -f $ObjType, $ObjId
        if ($MetricsAll.ContainsKey($k)) { $met = $MetricsAll[$k] }
    }
    $ctx = New-PlContext -ObjType $ObjType -ObjId $ObjId -TypeName (Get-AlTypeName $ObjType) -Lex $lex -M $met

    # построчная карта структуры
    $lineInfo = @{}
    foreach ($lr in $struct.Lines) { $lineInfo[$lr.LineNo] = $lr }
    $funcRel = @{}
    foreach ($fn in $struct.Functions) { $funcRel[$fn.Name] = $fn.Reliable }

    # переменные итерации по каждому циклу: X.NEXT в строке UNTIL (или NEXT под WITH)
    $loopIter = @{}   # StartLine -> HashSet имён
    foreach ($fn in $struct.Functions) {
        foreach ($lp in $fn.Loops) {
            $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            if ($lp.Kind -eq 'REPEAT' -and $lp.EndLine -ge 1 -and $lp.EndLine -le $lex.Count) {
                $uc = $lex[$lp.EndLine - 1].Clean
                foreach ($m in [regex]::Matches($uc, '(?<v>"[^"]+"|[A-Za-z_]\w*)\s*\.\s*NEXT(?![\w])')) {
                    [void]$set.Add($m.Groups['v'].Value.Trim('"'))
                }
                if ($set.Count -eq 0 -and $uc -match '(?<![\w".])NEXT(?![\w])') {
                    $li0 = $null
                    if ($lineInfo.ContainsKey($lp.EndLine)) { $li0 = $lineInfo[$lp.EndLine] }
                    if ($li0 -and $li0.WithVar) { [void]$set.Add($li0.WithVar.Trim('"')) }
                }
            }
            $loopIter[('{0}|{1}' -f $fn.Name, $lp.StartLine)] = $set
        }
    }

    # присваивания в теле функций (для оценки инвариантности аргументов GET)
    $assigns = @{}    # имя функции -> ArrayList {Line, Root}
    foreach ($fn in $struct.Functions) {
        $al = New-Object System.Collections.ArrayList
        for ($l = $fn.FirstLine; $l -le $fn.LastLine; $l++) {
            $c = $lex[$l - 1]
            if ($c.Kind -ne 'Code' -or -not $c.Clean) { continue }
            $ma = [regex]::Match($c.Clean, '^(?<root>"[^"]+"|[A-Za-z_]\w*)(?:\s*\.\s*(?:"[^"]+"|[A-Za-z_]\w*))*\s*(?::=|\+=|-=|\*=|/=)')
            if ($ma.Success) { [void]$al.Add(@{ Line = $l; Root = $ma.Groups['root'].Value.Trim('"') }) }
            # меняют значение переменной не только присваивания: методы записи и CLEAR
            foreach ($mv in [regex]::Matches($c.Clean,
                '(?<v>"[^"]+"|[A-Za-z_]\w*)\s*\.\s*(?:VALIDATE|TRANSFERFIELDS|INIT|COPY|GET|FIND|FINDSET|FINDFIRST|FINDLAST|NEXT)(?![\w])')) {
                [void]$al.Add(@{ Line = $l; Root = $mv.Groups['v'].Value.Trim('"') })
            }
            foreach ($mv in [regex]::Matches($c.Clean, '(?<![\w".])CLEAR\s*\(\s*(?<v>"[^"]+"|[A-Za-z_]\w*)')) {
                [void]$al.Add(@{ Line = $l; Root = $mv.Groups['v'].Value.Trim('"') })
            }
        }
        $assigns[$fn.Name] = $al
    }

    # внешние переменные для правила 115
    $extVars = @{}    # 'область|имя' -> краткий тип
    foreach ($s in $syms) {
        $t = $s.DataType
        if (-not $t) { continue }
        $flag = $false
        if ($t -match '^(Automation|OCX)') { $flag = $true }
        elseif ($t -match '^DotNet' -and $script:PlHeavyDotNetRx.IsMatch($t)) { $flag = $true }
        if ($flag) {
            $short = $t
            if ($short.Length -gt 40) { $short = $short.Substring(0, 38) + '..' }
            $extVars[($s.Scope + '|' + $s.Name)] = $short
        }
    }

    $reads     = @{}   # 'функция|перем' -> первая строка чтения
    $locks     = New-Object System.Collections.ArrayList   # {Func, Var, Line, VarInfo}
    $blobSeen  = @{}   # 'функция|перем|поле' -> первая строка CALCFIELDS BLOB
    $callSeen  = @{}   # правило 204: 'функция|выражение' -> ArrayList {Line, LoopStart}

    $filterRx = [regex]('(?<![\w".])(?:(?<var>"[^"]+"|[A-Za-z_][A-Za-z0-9_]*)\s*\.\s*)?' +
        '(?<op>SETRANGE|SETFILTER)(?![\w])\s*\(')
    $findLegacyRx = [regex]("(?<![\w"".])(?:(?<var>""[^""]+""|[A-Za-z_][A-Za-z0-9_]*)\s*\.\s*)?FIND\s*\(\s*''\s*\)")

    # ------------------------------------------------------------------ проход 1
    foreach ($row in $lex) {
        if ($row.Kind -ne 'Code' -or -not $row.Clean) { continue }
        $ln = $row.LineNo
        $li = $null
        if ($lineInfo.ContainsKey($ln)) { $li = $lineInfo[$ln] }
        $func = $row.FunctionName
        if ($li) { $func = $li.Function }
        # «В цикле» — это не только REPEAT/FOR/WHILE. Признак берётся из трёх
        # источников, от синтаксиса к факту:
        #   1) явный цикл в тексте;
        #   2) триггер, который платформа зовёт на каждой записи (см. выше);
        #   3) замер: строка, отработавшая сотни раз, находится в цикле — как бы
        #      он ни был устроен. Это самое сильное свидетельство из трёх, потому
        #      что оно не зависит от разбора кода.
        $inLoop = ($li -and $li.LoopDepth -ge 1)
        if (-not $inLoop -and (Test-PlImplicitLoop -ObjType $ObjType -Func $func)) { $inLoop = $true }
        if (-not $inLoop -and $met -and $met.Lines.ContainsKey($ln) -and
            $met.Lines[$ln].Hits -ge $LoopHits) { $inLoop = $true }
        $impRec = Get-PlImplicitRecord -ObjType $ObjType -Func $func
        $withVar = ''
        if ($li -and $li.WithVar) { $withVar = $li.WithVar.Trim('"') }
        $clean = $row.Clean
        $conf2 = 'средняя'
        if ($funcRel.ContainsKey($func) -and -not $funcRel[$func]) { $conf2 = 'низкая' }   # разбор функции не сошёлся

        # --- 101: COUNT в сравнении с 0/1 --------------------------------
        if ($clean.IndexOf('COUNT') -ge 0) {
            foreach ($m in $script:PlCountCmpRx.Matches($clean)) {
                $var = ''
                if ($m.Groups['var'].Success) { $var = $m.Groups['var'].Value.Trim('"') }
                elseif ($withVar) { $var = $withVar }
                elseif ($impRec) { $var = $impRec }
                else { continue }
                $vi = Resolve-PlVar $var $func $symIdx $ObjType $ObjId
                if ($m.Groups['var'].Success -and (-not $vi.Known -or -not $vi.IsRecord)) { continue }  # DotNet .Count и пр.
                if (-not $m.Groups['var'].Success -and $vi.Known -and -not $vi.IsRecord) { continue }
                if (Test-PlTempSkip $vi $met $ln) { continue }
                $op = $m.Groups['op'].Value; $n = $m.Groups['n'].Value
                $repl = ''
                if (($op -eq '=' -and $n -eq '0') -or ($op -eq '<=' -and $n -eq '0') -or ($op -eq '<' -and $n -eq '1')) { $repl = 'ISEMPTY' }
                elseif (($op -eq '<>' -and $n -eq '0') -or ($op -eq '>' -and $n -eq '0') -or ($op -eq '>=' -and $n -eq '1')) { $repl = 'NOT ISEMPTY' }
                if (-not $repl) { continue }
                Add-PlFinding -Ctx $ctx -Line $ln -Rule '101' -Func $func -Msg (
                    ('COUNT для проверки пустоты по {0}: замените на {1} — COUNT читает все строки по фильтру, ISEMPTY останавливается на первой' -f $var, $repl)) -Conf 'высокая'
            }
        }

        # --- 102: легаси FIND('-') / FIND('+') ---------------------------
        if ($clean.IndexOf("FIND(''") -ge 0) {
            $occ = @{}
            foreach ($m in $findLegacyRx.Matches($clean)) {
                $var = ''
                if ($m.Groups['var'].Success) { $var = $m.Groups['var'].Value.Trim('"') }
                elseif ($withVar) { $var = $withVar }
                elseif ($impRec) { $var = $impRec }
                $key = $var
                if (-not $occ.ContainsKey($key)) { $occ[$key] = 0 }
                $ix = $occ[$key]; $occ[$key] = $ix + 1
                $vi = Resolve-PlVar $var $func $symIdx $ObjType $ObjId
                if ($m.Groups['var'].Success -and (-not $vi.Known -or -not $vi.IsRecord)) { continue }
                if (-not $var) { continue }
                if (Test-PlTempSkip $vi $met $ln) { continue }
                $argVar = $var
                if (-not $m.Groups['var'].Success) { $argVar = '' }
                $lit = Get-PlLiteral (Get-PlCallArgText $rows $ln $argVar 'FIND' $ix)
                if ($null -eq $lit) { continue }
                if ($lit -eq '-') {
                    Add-PlFinding -Ctx $ctx -Line $ln -Rule '102' -Func $func -Msg (
                        ("легаси FIND('-') по {0}: для перебора — FINDSET, для проверки/первой записи — FINDFIRST (не тянут весь набор)" -f $var)) -Conf 'высокая'
                }
                elseif ($lit -eq '+') {
                    Add-PlFinding -Ctx $ctx -Line $ln -Rule '102' -Func $func -Msg (
                        ("легаси FIND('+') по {0}: замените на FINDLAST" -f $var)) -Conf 'высокая'
                }
            }
        }

        # --- 103/104: SETFILTER-паттерны и фильтр по FlowField -----------
        if ($clean.IndexOf('SETRANGE') -ge 0 -or $clean.IndexOf('SETFILTER') -ge 0) {
            $occ = @{}
            foreach ($m in $filterRx.Matches($clean)) {
                $op  = $m.Groups['op'].Value
                $var = ''
                if ($m.Groups['var'].Success) { $var = $m.Groups['var'].Value.Trim('"') }
                elseif ($withVar) { $var = $withVar }
                elseif ($impRec) { $var = $impRec }
                $key = $var + '|' + $op
                if (-not $occ.ContainsKey($key)) { $occ[$key] = 0 }
                $ix = $occ[$key]; $occ[$key] = $ix + 1
                if (-not $var) { continue }
                $vi = Resolve-PlVar $var $func $symIdx $ObjType $ObjId
                if ($m.Groups['var'].Success -and $vi.Known -and -not $vi.IsRecord) { continue }
                $argVar = $var
                if (-not $m.Groups['var'].Success) { $argVar = '' }
                $argTxt = Get-PlCallArgText $rows $ln $argVar $op $ix
                if ($null -eq $argTxt) { continue }
                $argList = Split-PlArgs $argTxt
                if ($argList.Count -lt 1) { continue }
                $fld = Get-PlFieldArg $argList[0]

                # 104 — фильтр по FlowField: точный факт метаданных
                if ($vi.IsRecord -and $vi.TableNo -gt 0 -and $fld) {
                    $fMeta = Get-KaField -TableId $vi.TableNo -Name $fld
                    if ($fMeta -and $fMeta.FieldClass -eq 'FlowField') {
                        $cf = $fMeta.CalcFormula
                        if ($cf -and $cf.Length -gt 60) { $cf = $cf.Substring(0, 58) + '..' }
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '104' -Func $func `
                            -Elem ('{0}."{1}"' -f $var, $fld) -Msg (
                            ('фильтр по FlowField {0}."{1}": платформа вычисляет поле для КАЖДОЙ строки-кандидата (CalcFormula: {2}); фильтруйте по базовым полям или сузьте выборку до фильтра' -f $var, $fld, $cf)) -Conf 'высокая'
                    }
                }

                # 103 — SETFILTER, который не ложится в индекс
                if ($op -eq 'SETFILTER' -and $argList.Count -ge 2 -and -not (Test-PlTempSkip $vi $met $ln)) {
                    $lit = Get-PlLiteral $argList[1]
                    if ($null -ne $lit) {
                        $lt = $lit.Trim()
                        $show = $lt
                        if ($show.Length -gt 30) { $show = $show.Substring(0, 28) + '..' }
                        if ($lt.StartsWith('<>')) {
                            Add-PlFinding -Ctx $ctx -Line $ln -Rule '103' -Func $func -Msg (
                                ("SETFILTER {0}, '{1}' по {2}: отрицание не ложится в индекс — SQL читает весь диапазон и отсеивает; сформулируйте позитивно (диапазон/набор значений) или убедитесь, что остальной фильтр узкий" -f $fld, $show, $var)) -Conf 'высокая' -Elem ('{0}."{1}"' -f $var, $fld)
                        }
                        elseif ($lt.StartsWith('*') -or $lt.StartsWith('@*')) {
                            Add-PlFinding -Ctx $ctx -Line $ln -Rule '103' -Func $func -Msg (
                                ("SETFILTER {0}, '{1}' по {2}: маска с '*' слева не использует индекс — поиск перебором; если достаточно поиска по началу строки, уберите ведущую '*'" -f $fld, $show, $var)) -Conf 'высокая' -Elem ('{0}."{1}"' -f $var, $fld)
                        }
                    }
                }
            }
        }

        # --- единый разбор операторов строки -----------------------------
        $occ = @{}
        foreach ($m in $script:PlLineOpRx.Matches($clean)) {
            $op  = $m.Groups['op'].Value
            $hasVar = $m.Groups['var'].Success
            $var = ''
            if ($hasVar) { $var = $m.Groups['var'].Value.Trim('"') }
            elseif ($withVar) { $var = $withVar }
            elseif ($impRec) { $var = $impRec }
            $key = $var + '|' + $op
            if (-not $occ.ContainsKey($key)) { $occ[$key] = 0 }
            $ix = $occ[$key]; $occ[$key] = $ix + 1
            $argVar = $var
            if (-not $hasVar) { $argVar = '' }
            $vi = Resolve-PlVar $var $func $symIdx $ObjType $ObjId

            switch ($op) {
                'COMMIT' {
                    if (-not $hasVar -and $inLoop) {
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '114' -Func $func -Msg (
                            'COMMIT в цикле: фиксация транзакции на каждой итерации — запись журнала SQL и сброс блокировок; вынести за цикл или коммитить пакетами') -Conf 'высокая'
                    }
                }
                'SLEEP' {
                    if (-not $hasVar -and $inLoop) {
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '115' -Func $func -Msg (
                            'SLEEP в цикле — искусственная пауза умножается на число итераций') -Conf 'высокая'
                    }
                }
                'SHELL' {
                    if (-not $hasVar -and $inLoop) {
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '115' -Func $func -Msg (
                            'SHELL в цикле — запуск внешнего процесса на каждой итерации') -Conf 'высокая'
                    }
                }
                'CALCFIELDS' {
                    if (-not $vi.Known -and $hasVar) { $cfConf = 'низкая' } else { $cfConf = $conf2 }
                    if ($vi.Known -and -not $vi.IsRecord) { break }
                    $argTxt = Get-PlCallArgText $rows $ln $argVar 'CALCFIELDS' $ix
                    $flds = @()
                    if ($null -ne $argTxt) {
                        $cfArgs = Split-PlArgs $argTxt      # через переменную: см. контракт Split-PlArgs
                        $flds = @($cfArgs | ForEach-Object { Get-PlFieldArg $_ } | Where-Object { $_ })
                    }
                    if ($inLoop) {
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '110' -Func $func -Msg (
                            ('CALCFIELDS({0}) по {1} в цикле — отдельные запросы на каждой итерации; SETAUTOCALCFIELDS перед выборкой или вынести из цикла' -f (($flds -join ', ')), $var)) -Conf $cfConf
                    }
                    # 118 — чтение BLOB
                    if ($vi.IsRecord -and $vi.TableNo -gt 0 -and -not (Test-PlTempSkip $vi $met $ln)) {
                        foreach ($f in $flds) {
                            $fMeta = Get-KaField -TableId $vi.TableNo -Name $f
                            if (-not $fMeta -or $fMeta.DataType -ne 'BLOB') { continue }
                            $bk = '{0}|{1}|{2}' -f $func, $var, $f
                            if ($inLoop) {
                                Add-PlFinding -Ctx $ctx -Line $ln -Rule '118' -Func $func -Elem ('{0}."{1}"' -f $var, $f) -Msg (
                                    ('чтение BLOB {0}."{1}" в цикле: содержимое тянется с сервера на каждой итерации — прочитайте один раз и кэшируйте' -f $var, $f)) -Conf 'высокая'
                            }
                            elseif ($blobSeen.ContainsKey($bk)) {
                                Add-PlFinding -Ctx $ctx -Line $ln -Rule '118' -Func $func -Elem ('{0}."{1}"' -f $var, $f) -Msg (
                                    ('повторное чтение BLOB {0}."{1}": уже прочитан на строке {2} — переиспользуйте содержимое' -f $var, $f, $blobSeen[$bk])) -Conf 'высокая'
                            }
                            else { $blobSeen[$bk] = $ln }
                        }
                    }
                }
                'GET' {
                    if (-not $var) { break }
                    if (-not $vi.Known -or -not $vi.IsRecord) { break }   # DotNet .Get и пр.
                    $rk = '{0}|{1}' -f $func, $var
                    if (-not $reads.ContainsKey($rk)) { $reads[$rk] = $ln }
                    if (-not $inLoop) { break }
                    if (Test-PlTempSkip $vi $met $ln) { break }
                    # инвариантность аргументов относительно объемлющих циклов
                    $argTxt = Get-PlCallArgText $rows $ln $argVar 'GET' $ix
                    $variant = $false
                    if ($null -ne $argTxt -and $argTxt.Trim()) {
                        $argRoots = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                        $prevDot = $false
                        foreach ($mi in [regex]::Matches($argTxt, '("[^"]+"|[A-Za-z_]\w*)|(\S)')) {
                            $tok = $mi.Value
                            if ($tok -eq '.') { $prevDot = $true; continue }
                            if ($tok -match '^["A-Za-z_]') {
                                if (-not $prevDot) { [void]$argRoots.Add($tok.Trim('"')) }
                                $prevDot = $false
                            }
                            else { $prevDot = $false }
                        }
                        $fnRec = $null
                        foreach ($fn0 in $struct.Functions) { if ($fn0.Name -eq $func) { $fnRec = $fn0; break } }
                        if ($fnRec) {
                            foreach ($lp in $fnRec.Loops) {
                                if ($ln -lt $lp.StartLine -or $ln -gt $lp.EndLine) { continue }
                                $itKey = '{0}|{1}' -f $func, $lp.StartLine
                                if ($loopIter.ContainsKey($itKey)) {
                                    foreach ($iv in $loopIter[$itKey]) { if ($argRoots.Contains($iv)) { $variant = $true; break } }
                                }
                                if ($variant) { break }
                                if ($assigns.ContainsKey($func)) {
                                    foreach ($as in $assigns[$func]) {
                                        if ($as.Line -ge $lp.StartLine -and $as.Line -le $lp.EndLine -and $argRoots.Contains($as.Root)) { $variant = $true; break }
                                    }
                                }
                                if ($variant) { break }
                            }
                        }
                        else { $variant = $true }
                    }
                    else { $variant = $true }   # аргументы не разобраны — не утверждаем инвариантность
                    if (-not $variant) {
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '111' -Func $func -Msg (
                            ('GET по {0} в цикле: аргументы не меняются между итерациями — значение инвариантно, вынести до цикла' -f $var)) -Conf 'высокая'
                    }
                    else {
                        Add-PlFinding -Ctx $ctx -Line $ln -Rule '111' -Func $func -Msg (
                            ('GET по {0} в цикле (N+1): точечный запрос на каждой итерации — кэшировать (временная таблица/словарь) или получить данные одной выборкой' -f $var)) -Conf 'низкая'
                    }
                }
                'LOCKTABLE' {
                    if (-not $var) { break }
                    [void]$locks.Add(@{ Func = $func; Var = $var; Line = $ln; VarInfo = $vi })
                }
                { $_ -eq 'MODIFY' -or $_ -eq 'INSERT' -or $_ -eq 'DELETE' } {
                    if (-not $inLoop -or -not $var) { break }
                    if (-not $vi.Known -or -not $vi.IsRecord) { break }
                    if (Test-PlTempSkip $vi $met $ln) { break }
                    $isIter = $false
                    if ($li -and $li.LoopStart -gt 0) {
                        $itKey = '{0}|{1}' -f $func, $li.LoopStart
                        if ($loopIter.ContainsKey($itKey) -and $loopIter[$itKey].Contains($var)) { $isIter = $true }
                    }
                    $msg = ''
                    if ($op -eq 'DELETE' -and $isIter) {
                        $msg = ('DELETE перебираемой {0} построчно: если прочая логика в цикле не нужна — одним DELETEALL по тем же фильтрам' -f $var)
                    }
                    elseif ($op -eq 'MODIFY' -and $isIter) {
                        $msg = ('MODIFY перебираемой {0} на каждой итерации: если правка сводится к константе по фильтру — MODIFYALL (внимание: триггеры не выполняет)' -f $var)
                    }
                    elseif ($op -eq 'INSERT') {
                        $msg = ('построчная вставка {0} в цикле — запрос на строку; часто неизбежно, судите по замеру строки' -f $var)
                    }
                    else {
                        $msg = ('построчный {0} по {1} в цикле — запрос на строку; проверьте, нельзя ли пакетно (MODIFYALL/DELETEALL) либо реже' -f $op, $var)
                    }
                    $cf = 'средняя'
                    if ($op -eq 'INSERT') { $cf = 'низкая' }
                    Add-PlFinding -Ctx $ctx -Line $ln -Rule '113' -Func $func -Msg $msg -Conf $cf
                }
                { $_ -eq 'FIND' -or $_ -eq 'FINDSET' -or $_ -eq 'FINDFIRST' -or $_ -eq 'FINDLAST' -or $_ -eq 'COUNT' } {
                    if ($var -and $vi.Known -and $vi.IsRecord) {
                        $rk = '{0}|{1}' -f $func, $var
                        if (-not $reads.ContainsKey($rk)) { $reads[$rk] = $ln }
                    }
                }
            }
        }

        # --- 115: вызовы внешних объектов (Automation/OCX/тяжёлый DotNet) в цикле
        if ($inLoop -and $extVars.Count -gt 0) {
            foreach ($ek in $extVars.Keys) {
                $p = $ek.Split('|')
                if ($p[0] -and $p[0] -ne $func) { continue }
                $nm = $p[1]
                $pat = ('(?<![\w".])(?:"{0}"|{0})\s*\.\s*[A-Za-z_]' -f [regex]::Escape($nm))
                if ([regex]::IsMatch($clean, $pat)) {
                    Add-PlFinding -Ctx $ctx -Line $ln -Rule '115' -Func $func -Elem $nm -Msg (
                        ('внешний вызов {0} ({1}) в цикле — межпроцессный/сетевой вызов на каждой итерации; пакетируйте или выносите из цикла' -f $nm, $extVars[$ek])) -Conf 'средняя'
                }
            }
        }

        # --- 204: сбор одинаковых вызовов-выражений ----------------------
        # Очередь с вложенностью: PlCallRx съедает скобки целиком, и вызов внутри
        # аргументов (MarkingType.GET(GetMarkTypeCode(..))) иначе не виден.
        # Самостоятельный оператор-вызов (вся строка = Вызов(...);) пропускается:
        # он исполняется ради побочного эффекта, повтор почти всегда намеренный.
        # Сканируем СЫРУЮ строку без //-комментария, а не Clean: лексер заменяет
        # литералы на '' и разные вызовы GetVar('ITEM')/GetVar('LOT') становились
        # «одинаковыми». Имена в кавычках содержат скобки и пробелы
        # ("Quantity (Base)") и без предобработки распознаются как «вызовы»;
        # скобки внутри кавычек глушим, сохраняя различимость имён.
        $scan204 = Remove-PlLineComment $rows[$ln - 1].Text
        if ($scan204.IndexOf('"') -ge 0) {
            $scan204 = [regex]::Replace($scan204, '"[^"]*"',
                { param($mm) '"' + ($mm.Value.Trim('"') -replace '[()\s]', '_') + '"' })
        }
        $queue204 = New-Object System.Collections.ArrayList
        [void]$queue204.Add(@{ Text = $scan204; Nested = $false })
        $qi = 0
        while ($qi -lt $queue204.Count) {
            $q0 = $queue204[$qi]; $qi++
            if ($qi -gt 24) { break }
            foreach ($m in $script:PlCallRx.Matches($q0.Text)) {
                $args0 = $m.Groups['args'].Value.Trim()
                if ($args0) { [void]$queue204.Add(@{ Text = $args0; Nested = $true }) }
                if (-not $args0) { continue }
                $head = $m.Groups['head'].Value
                $dotIx = $head.LastIndexOf('.')
                $fname = $head
                if ($dotIx -ge 0) { $fname = $head.Substring($dotIx + 1) }
                $fname = $fname.Trim().Trim('"')
                if (-not $fname -or $script:PlBuiltins.Contains($fname)) { continue }
                if (-not $q0.Nested) {
                    # оператор-вызов: до матча только отступ, после него только ';'
                    $tail204 = $q0.Text.Substring($m.Index + $m.Length).Trim()
                    $lead204 = $q0.Text.Substring(0, $m.Index).Trim()
                    if ($lead204 -eq '' -and ($tail204 -eq '' -or $tail204 -eq ';')) { continue }
                }
                $expr = [regex]::Replace($m.Value, '\s+', '').ToUpperInvariant()
                # корневые идентификаторы аргументов: их присваивание между двумя
                # вхождениями делает «повтор» честным пересчётом
                $roots204 = New-Object System.Collections.ArrayList
                $prevDot204 = $false
                foreach ($mi in [regex]::Matches($m.Groups['args'].Value, '("[^"]+"|[A-Za-z_]\w*)|(\S)')) {
                    $tok = $mi.Value
                    if ($tok -eq '.') { $prevDot204 = $true; continue }
                    if ($tok -match '^["A-Za-z_]') {
                        if (-not $prevDot204) { [void]$roots204.Add($tok.Trim('"')) }
                        $prevDot204 = $false
                    }
                    else { $prevDot204 = $false }
                }
                $ck = $func + '|' + $expr
                if (-not $callSeen.ContainsKey($ck)) { $callSeen[$ck] = New-Object System.Collections.ArrayList }
                $ls = 0
                if ($li -and $li.LoopStart -gt 0) { $ls = $li.LoopStart }
                [void]$callSeen[$ck].Add(@{ Line = $ln; LoopStart = $ls; Show = $m.Value; Roots = $roots204 })
            }
        }
    }

    # ------------------------------------------------------------- 119: поздний LOCKTABLE
    # чтения пополняются потребителями цепочек (ISEMPTY и пр. в проходе 1 не ловятся)
    foreach ($c in $chains) {
        $rk = '{0}|{1}' -f $c.Function, $c.Variable
        if (-not $reads.ContainsKey($rk) -or $c.ConsumerLine -lt $reads[$rk]) { $reads[$rk] = $c.ConsumerLine }
    }
    foreach ($lk in $locks) {
        if (-not $lk.VarInfo.Known -or -not $lk.VarInfo.IsRecord) { continue }
        if (Test-PlTempSkip $lk.VarInfo $met $lk.Line) { continue }
        $rk = '{0}|{1}' -f $lk.Func, $lk.Var
        if ($reads.ContainsKey($rk) -and $reads[$rk] -lt $lk.Line) {
            Add-PlFinding -Ctx $ctx -Line $lk.Line -Rule '119' -Func $lk.Func -Msg (
                ('LOCKTABLE после чтения {0} (первое чтение — строка {1}): прочитанное без блокировки будет перечитано с UPDLOCK; вызвать LOCKTABLE до первого чтения или читать FINDSET(TRUE)' -f $lk.Var, $reads[$rk])) -Conf 'средняя'
        }
    }


    # ------------------------------------------------------------- правила по цепочкам
    $consumerOps = @('FINDSET', 'FINDFIRST', 'FINDLAST', 'FIND', 'COUNT', 'ISEMPTY', 'CALCSUMS', 'DELETEALL', 'MODIFYALL')
    $recOnlyOps  = @('FINDSET', 'FINDFIRST', 'FINDLAST', 'CALCSUMS', 'ISEMPTY', 'DELETEALL', 'MODIFYALL')

    # 112: запрос внутри цикла (N+1); инвариантный — вынести
    foreach ($c in $chains) {
        if ($consumerOps -notcontains $c.ConsumerOp) { continue }
        $li = $null
        if ($lineInfo.ContainsKey($c.ConsumerLine)) { $li = $lineInfo[$c.ConsumerLine] }
        if (-not $li -or $li.LoopDepth -lt 1) { continue }
        $conf = 'средняя'
        if (-not $c.Resolved) {
            if ($recOnlyOps -notcontains $c.ConsumerOp) { continue }   # COUNT/FIND без типа — не рискуем
            $conf = 'низкая'
        }
        elseif ($c.Temporary) {
            $viTmp = [pscustomobject]@{ Known = $true; IsRecord = $true; TableNo = $c.TableNo; TableName = $c.TableName; Temporary = $true }
            if (Test-PlTempSkip $viTmp $met $c.ConsumerLine) { continue }
        }
        # действующий фильтр, выставленный внутри цикла, = запрос зависит от итерации
        # (именно ActiveFilterLines: свежие мог «съесть» предыдущий потребитель
        # той же итерации — ISEMPTY перед CALCSUMS, — а зависимость осталась)
        $freshInLoop = $false
        foreach ($fl in $c.ActiveFilterLines) {
            if ($fl.LineNo -ge $li.LoopStart -and $fl.LineNo -le $c.ConsumerLine) { $freshInLoop = $true; break }
        }
        $isIter = $false
        $itKey = '{0}|{1}' -f $c.Function, $li.LoopStart
        if ($loopIter.ContainsKey($itKey) -and $loopIter[$itKey].Contains($c.Variable)) { $isIter = $true }
        if ($isIter) { continue }   # сам перебираемый набор — не запрос в цикле
        $tblTxt = ''
        if ($c.TableName) { $tblTxt = ' (' + $c.TableName + ')' }
        if (-not $freshInLoop) {
            Add-PlFinding -Ctx $ctx -Line $c.ConsumerLine -Rule '112' -Func $c.Function -Elem ('{0}{1}' -f $c.Variable, $tblTxt) -Msg (
                ('{0} по {1}{2} в цикле, фильтры внутри итерации не меняются — результат одинаков на каждом витке: вынести из цикла' -f $c.ConsumerOp, $c.Variable, $tblTxt)) -Conf 'высокая'
        }
        else {
            Add-PlFinding -Ctx $ctx -Line $c.ConsumerLine -Rule '112' -Func $c.Function -Elem ('{0}{1}' -f $c.Variable, $tblTxt) -Msg (
                ('запрос в цикле (N+1): {0} по {1}{2} на каждой итерации; вложенный перебор типовой — судить по замеру, альтернатива: один запрос диапазоном или временная таблица' -f $c.ConsumerOp, $c.Variable, $tblTxt)) -Conf 'низкая'
        }
    }

    # 116: повторный одинаковый запрос без смены фильтров
    $byVar = @{}
    foreach ($c in $chains) {
        if ($consumerOps -notcontains $c.ConsumerOp) { continue }
        $k = '{0}|{1}' -f $c.Function, $c.Variable
        if (-not $byVar.ContainsKey($k)) { $byVar[$k] = New-Object System.Collections.ArrayList }
        [void]$byVar[$k].Add($c)
    }
    foreach ($k in $byVar.Keys) {
        $list = @($byVar[$k] | Sort-Object ConsumerLine)
        for ($i = 1; $i -lt $list.Count; $i++) {
            $cur = $list[$i]; $prev = $list[$i - 1]
            if ($cur.ConsumerOp -ne $prev.ConsumerOp) { continue }
            if (@($cur.FilterLines).Count -gt 0) { continue }
            # FIND — точечный re-read (FIND('=')), почти всегда намеренный refresh;
            # DELETEALL/MODIFYALL меняют состояние — их повтор не «тот же результат»
            if ($cur.ConsumerOp -in @('FIND', 'DELETEALL', 'MODIFYALL')) { continue }
            # CALCSUMS с разными полями сумм — разные запросы (ветки IF/ELSE)
            if ($cur.ConsumerOp -eq 'CALCSUMS') {
                $av = $cur.Variable
                if ($cur.ViaWith) { $av = '' }
                $a1 = Get-PlCallArgText $rows $prev.ConsumerLine $av 'CALCSUMS' 0
                $a2 = Get-PlCallArgText $rows $cur.ConsumerLine  $av 'CALCSUMS' 0
                $n1 = ''; $n2 = ''
                if ($null -ne $a1) { $n1 = ([regex]::Replace($a1, '\s+', '')).ToUpperInvariant() }
                if ($null -ne $a2) { $n2 = ([regex]::Replace($a2, '\s+', '')).ToUpperInvariant() }
                if ($n1 -ne $n2) { continue }
            }
            if ($cur.Temporary) {
                $viTmp = [pscustomobject]@{ Known = $true; IsRecord = $true; TableNo = $cur.TableNo; TableName = $cur.TableName; Temporary = $true }
                if (Test-PlTempSkip $viTmp $met $cur.ConsumerLine) { continue }
            }
            if (-not $cur.Resolved -and $recOnlyOps -notcontains $cur.ConsumerOp) { continue }
            Add-PlFinding -Ctx $ctx -Line $cur.ConsumerLine -Rule '116' -Func $cur.Function -Elem $cur.Variable -Msg (
                ('повторный {0} по {1} с теми же фильтрами — тот же запрос уже выполнен на строке {2}; сохраните результат (запись/флаг) и переиспользуйте' -f $cur.ConsumerOp, $cur.Variable, $prev.ConsumerLine)) -Conf 'средняя'
        }
    }

    # 301: советник по ключам (делегирование Lib-KeyAdvisor)
    foreach ($c in $chains) {
        if ($consumerOps -notcontains $c.ConsumerOp) { continue }
        if (-not $c.Resolved -or $c.TableNo -le 0) { continue }
        if ($c.Temporary) {
            $viTmp = [pscustomobject]@{ Known = $true; IsRecord = $true; TableNo = $c.TableNo; TableName = $c.TableName; Temporary = $true }
            if (Test-PlTempSkip $viTmp $met $c.ConsumerLine) { continue }
        }
        $cls = Get-TableClass -TableId $c.TableNo
        if ($cls.Silent) { continue }
        if (@($c.ActiveFilterLines).Count -eq 0) { continue }   # без фильтров это полный перебор — тема правила 112

        # восстановление состава фильтра из строк SETRANGE/SETFILTER/SETCURRENTKEY
        $fldCls  = @{}       # поле(lower) -> Eq | Range
        $fldOrd  = New-Object System.Collections.ArrayList
        $orderBy = @()
        $hasOther = $false
        $occA = @{}
        foreach ($fl in $c.ActiveFilterLines) {
            $ok2 = '{0}|{1}' -f $fl.LineNo, $fl.Op
            if (-not $occA.ContainsKey($ok2)) { $occA[$ok2] = 0 }
            $ix2 = $occA[$ok2]; $occA[$ok2] = $ix2 + 1
            if ($fl.Op -eq 'RESET') { continue }
            $argVar = $c.Variable
            if ($c.ViaWith) { $argVar = '' }
            $argTxt = Get-PlCallArgText $rows $fl.LineNo $argVar $fl.Op $ix2
            if ($null -eq $argTxt) {
                if ($c.ViaWith) { $argTxt = Get-PlCallArgText $rows $fl.LineNo $c.Variable $fl.Op $ix2 }
                if ($null -eq $argTxt) { $hasOther = $true; continue }
            }
            $argList = Split-PlArgs $argTxt
            if ($fl.Op -eq 'SETCURRENTKEY') {
                $orderBy = @($argList | ForEach-Object { Get-PlFieldArg $_ } | Where-Object { $_ })
                continue
            }
            if ($argList.Count -lt 1) { continue }
            $fld = Get-PlFieldArg $argList[0]
            if (-not $fld) { continue }
            $fLow = $fld.ToLowerInvariant()
            if ($fl.Op -eq 'SETRANGE') {
                if ($argList.Count -eq 1) { $fldCls.Remove($fLow); continue }        # снятие фильтра
                $cls2 = 'Eq'
                if ($argList.Count -ge 3 -and $argList[1] -ne $argList[2]) { $cls2 = 'Range' }
                if (-not $fldCls.ContainsKey($fLow)) { [void]$fldOrd.Add(@{ Low = $fLow; Name = $fld }) }
                $fldCls[$fLow] = $cls2
            }
            else {   # SETFILTER
                if ($argList.Count -lt 2) { continue }
                $lit = Get-PlLiteral $argList[1]
                if ($null -eq $lit) { $hasOther = $true; continue }
                $lt = $lit.Trim()
                $cls2 = ''
                if     ($lt.StartsWith('<>') -or $lt.StartsWith('*') -or $lt.StartsWith('@*')) { $hasOther = $true }
                elseif ($lt.IndexOf('..') -ge 0)  { $cls2 = 'Range' }
                elseif ($lt.IndexOf('|') -ge 0)   { $cls2 = 'Eq' }
                elseif ($lt -match '^(>=|<=|>|<)') { $cls2 = 'Range' }
                elseif ($lt.IndexOf('*') -ge 0)   { $cls2 = 'Range' }   # маска-префикс
                elseif ($lt.IndexOf('&') -ge 0)   { $cls2 = 'Range' }
                else                               { $cls2 = 'Eq' }
                if ($cls2) {
                    if (-not $fldCls.ContainsKey($fLow)) { [void]$fldOrd.Add(@{ Low = $fLow; Name = $fld }) }
                    $fldCls[$fLow] = $cls2
                }
            }
        }
        $eq = @(); $rg = @()
        foreach ($fo in $fldOrd) {
            if (-not $fldCls.ContainsKey($fo.Low)) { continue }
            if ($fldCls[$fo.Low] -eq 'Eq') { $eq += $fo.Name } else { $rg += $fo.Name }
        }
        if ($eq.Count -eq 0 -and $rg.Count -eq 0) { continue }   # нечего сверять (частая причина — hasOther)

        $cov = Test-KeyCoverage -TableId $c.TableNo -Equality $eq -Range $rg -OrderBy $orderBy -HasOtherPredicates:$hasOther
        if ($cov.Verdict -notin @('NewKey', 'ExtendKey', 'KeyDisabled', 'KeyNotMaintained')) { continue }

        $act = 'Действие'
        if (-not $cls.Advise) { $act = 'Нужно, но не своими силами (диапазон вне лицензии)' }
        $tail = ''
        if ($cls.Class -eq 'Standard')  { $tail = '; таблица стандартная — правка уйдёт в upgrade-контур' }
        if ($cls.Class -eq 'ClientCore') { $tail = '; ядро отраслевого решения 10M+ — передать замер владельцу подсистемы' }

        $msg = ''
        switch ($cov.Verdict) {
            'NewKey' {
                $prop = New-KeyProposal -TableId $c.TableNo -Equality $eq -Range $rg -Coverage $cov
                if (@($prop.Fields).Count -eq 0) { break }
                $msg = ('{0} по {1} ({2}): подходящего ключа нет (равенство: {3}; диапазон: {4}). {5}: новый ключ "{6}". Цена: +1 индекс на каждую запись, изменение ключей = синхронизация схемы{7}' -f `
                    $c.ConsumerOp, $c.Variable, $c.TableName, ($eq -join ', '), ($rg -join ', '), $act, (@($prop.Fields) -join ','), $tail)
            }
            'ExtendKey' {
                $prop = New-KeyProposal -TableId $c.TableNo -Equality $eq -Range $rg -Coverage $cov
                # Дописывать нечего - и находки нет. У ветки NewKey такой выход есть,
                # у этой не было, и совет печатался с пустым перечнем: «дописать В
                # КОНЕЦ ключа 16:  (вставка в середину сдвинет порядок обхода всем)».
                # Пусто выходит штатно: все годные поля уже лежат в самом ключе.
                if (@($prop.AddedFields).Count -eq 0) { break }
                $msg = ('{0} по {1} ({2}): ключ {3} ({4}) целиком лежит в фильтре. {5}: дописать В КОНЕЦ ключа {3}: {6} (вставка в середину сдвинет порядок обхода всем){7}' -f `
                    $c.ConsumerOp, $c.Variable, $c.TableName, $cov.BestKeyNo, (@($cov.BestKey.Fields) -join ', '), $act, (@($prop.AddedFields) -join ', '), $tail)
            }
            'KeyDisabled' {
                $msg = ('{0} по {1} ({2}): ключ {3} нужного состава ВЫКЛЮЧЕН (Enabled=No). {4}: включить ключ {3} — состав уже подходит, нумерация ключей сквозная, остальные не сдвинутся{5}' -f `
                    $c.ConsumerOp, $c.Variable, $c.TableName, $cov.BestKeyNo, $act, $tail)
            }
            'KeyNotMaintained' {
                $msg = ('{0} по {1} ({2}): ключ {3} есть, но MaintainSQLIndex=No — физического индекса нет, SETCURRENTKEY ускорения не даст. {4}: MaintainSQLIndex=Yes{5}' -f `
                    $c.ConsumerOp, $c.Variable, $c.TableName, $cov.BestKeyNo, $act, $tail)
            }
        }
        if (-not $msg) { continue }

        # CALCSUMS: сумма без SumIndexFields считается перебором
        if ($c.ConsumerOp -eq 'CALCSUMS') {
            $argVar = $c.Variable
            if ($c.ViaWith) { $argVar = '' }
            $sumTxt = Get-PlCallArgText $rows $c.ConsumerLine $argVar 'CALCSUMS' 0
            if ($null -ne $sumTxt) {
                $sumArgs = Split-PlArgs $sumTxt            # через переменную: см. контракт Split-PlArgs
                foreach ($sf in @($sumArgs | ForEach-Object { Get-PlFieldArg $_ } | Where-Object { $_ })) {
                    $sLow = (ConvertTo-KaSqlName $sf).ToLowerInvariant()
                    $found = $false
                    foreach ($kk in (Get-KaTableKeys $c.TableNo)) { if ($kk.SumLower -contains $sLow) { $found = $true; break } }
                    if (-not $found) {
                        $msg += ('; CALCSUMS "{0}": поле не объявлено в SumIndexFields ни одного ключа — сумма считается перебором строк (SumIndexFields добавлять только по замеру)' -f $sf)
                    }
                }
            }
        }
        Add-PlFinding -Ctx $ctx -Line $c.ConsumerLine -Rule '301' -Func $c.Function `
            -Elem ('{0} ({1})' -f $c.Variable, $c.TableName) -Msg $msg -Conf 'средняя'
    }

    # ------------------------------------------------------------- 204: повторные выражения
    foreach ($ck in $callSeen.Keys) {
        $list = $callSeen[$ck]
        if ($list.Count -lt 2) { continue }
        $fname = $ck.Substring(0, $ck.IndexOf('|'))
        $sorted = @($list | Sort-Object { $_.Line })
        $fnAssigns = $null
        if ($assigns.ContainsKey($fname)) { $fnAssigns = $assigns[$fname] }
        # одна находка на группу «функция + выражение»: повторов бывают сотни
        # (S(RowNo,0) в выгрузке Excel), и по-находке-на-повтор топит отчёт
        $repeats = New-Object System.Collections.ArrayList
        $firstLine = $sorted[0].Line
        $prev = $sorted[0]
        for ($i = 1; $i -lt $sorted.Count; $i++) {
            $cur = $sorted[$i]
            $near = (($cur.Line - $prev.Line) -le 50) -or ($cur.LoopStart -gt 0 -and $cur.LoopStart -eq $prev.LoopStart)
            # корень аргумента присваивается между вхождениями — это пересчёт, не повтор
            if ($near -and $fnAssigns -and $cur.Roots.Count -gt 0) {
                foreach ($as in $fnAssigns) {
                    if ($as.Line -le $prev.Line -or $as.Line -gt $cur.Line) { continue }
                    if ($cur.Roots.Contains($as.Root)) { $near = $false; break }
                }
            }
            if ($near -and $cur.Line -ne $prev.Line) { [void]$repeats.Add($cur) }
            $prev = $cur
        }
        if ($repeats.Count -gt 0) {
            $first = $repeats[0]
            $conf = 'низкая'
            if ($met -and $met.Lines.ContainsKey($first.Line) -and $met.Lines.ContainsKey($firstLine) -and
                $met.Lines[$first.Line].Hits -gt 0 -and $met.Lines[$firstLine].Hits -gt 0) { $conf = 'средняя' }
            $show = $first.Show
            if ($show.Length -gt 50) { $show = $show.Substring(0, 48) + '..' }
            $where = @($repeats | ForEach-Object { $_.Line } | Select-Object -First 6) -join ', '
            $more = ''
            if ($repeats.Count -gt 6) { $more = ' и ещё ' + ($repeats.Count - 6) }
            Add-PlFinding -Ctx $ctx -Line $first.Line -Rule '204' -Func $fname -Msg (
                ('одинаковое выражение {0} вычислено повторно (впервые — строка {1}; повторы: {2}{3}) — сохраните результат в переменную (проверьте, что точки исполняются и аргументы между ними не меняются)' -f $show, $firstLine, $where, $more)) -Conf $conf
        }
    }

    # ------------------------------------------------------------- 201/202/203: только с метриками
    if ($met) {
        foreach ($fn in $struct.Functions) {
            $fs = 0.0
            if ($met.FuncSelf.ContainsKey($fn.Name)) { $fs = $met.FuncSelf[$fn.Name] }
            if ($fs -le 0) { continue }
            $pct = 0.0
            if ($met.TotalSelf -gt 0) { $pct = 100.0 * $fs / $met.TotalSelf }
            if (($fs -lt $HotFuncMs) -and ($pct -lt $HotFuncPct)) { continue }
            $fq = 0.0
            if ($met.FuncSql.ContainsKey($fn.Name)) { $fq = $met.FuncSql[$fn.Name] }
            $mh = 0L
            if ($met.FuncHits.ContainsKey($fn.Name)) { $mh = $met.FuncHits[$fn.Name] }
            Add-PlFinding -Ctx $ctx -Line $fn.HeaderLine -Rule '201' -Func $fn.Name -Severity 'High' -Conf 'точная' -GainMs $fs -Msg (
                ('горячая функция: {0:0.#} мс собственного времени ({1:0.#} % объекта), из них SQL {2:0.#} мс; макс. попаданий строки {3}' -f $fs, $pct, $fq, $mh))
            $sqlPct = 100.0 * $fq / $fs
            $dmsg = ''
            if     ($sqlPct -ge 60) { $dmsg = ('функция упирается в SQL ({0:0.} % времени в запросах): сначала ключи и фильтры (правила 103/104/301), микрооптимизации C/AL вторичны' -f $sqlPct) }
            elseif ($sqlPct -le 20) { $dmsg = ('функция упирается в C/AL (SQL лишь {0:0.} %): выигрыш в алгоритме и строках (110-116, 204), ключи не помогут' -f $sqlPct) }
            else                    { $dmsg = ('смешанный профиль: SQL {0:0.} % / C/AL {1:0.} % — смотреть и ключи, и код' -f $sqlPct, (100 - $sqlPct)) }
            Add-PlFinding -Ctx $ctx -Line $fn.HeaderLine -Rule '203' -Func $fn.Name -Severity 'Medium' -Conf 'точная' -GainMs 0 -Msg $dmsg
        }
        foreach ($fn in $struct.Functions) {
            foreach ($lp in $fn.Loops) {
                $iter = 0L; $selfSum = 0.0
                for ($l = $lp.StartLine; $l -le $lp.EndLine; $l++) {
                    if (-not $met.Lines.ContainsKey($l)) { continue }
                    $r = $met.Lines[$l]
                    if ($r.Hits -gt $iter) { $iter = $r.Hits }
                    $selfSum += $r.SelfMs
                }
                if ($iter -lt $LoopHits) { continue }
                $sev = 'Medium'
                if ($selfSum -ge $HotFuncMs) { $sev = 'High' }
                Add-PlFinding -Ctx $ctx -Line $lp.StartLine -Rule '202' -Func $fn.Name -Severity $sev -Conf 'точная' -GainMs $selfSum -Msg (
                    ('цикл {0} (строки {1}..{2}): ~{3} итераций за прогон, внутри набрано {4:0.#} мс — каждая находка в теле умножается на число итераций' -f $lp.Kind, $lp.StartLine, $lp.EndLine, $iter, $selfSum))
            }
        }
    }

    return ,$ctx.Findings
}


# ---------------------------------------------------------------------------
# Главный ход
# ---------------------------------------------------------------------------

$script:PlRuleNames = @{
    '101' = 'COUNT вместо ISEMPTY';      '102' = "легаси FIND('-')/FIND('+')"
    '103' = 'SETFILTER мимо индекса';    '104' = 'фильтр по FlowField'
    '110' = 'CALCFIELDS в цикле';        '111' = 'GET в цикле'
    '112' = 'запрос в цикле (N+1)';      '113' = 'построчная запись в цикле'
    '114' = 'COMMIT в цикле';            '115' = 'внешний вызов в цикле'
    '116' = 'повторный запрос';          '118' = 'повторное чтение BLOB'
    '119' = 'поздний LOCKTABLE';         '201' = 'горячая функция'
    '202' = 'кардинальность цикла';      '203' = 'диагноз SQL против C/AL'
    '204' = 'повторное выражение';       '301' = 'ключи (KeyAdvisor)'
}

$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
$srcRoot = Get-AlSourceRoot $SourceRoot
$baseDir = Get-AlBaseRoot $BaseRoot
Initialize-KeyAdvisor

$metricsAll = $null
if ($MetricsFile) {
    if (-not [System.IO.Path]::IsPathRooted($MetricsFile)) { $MetricsFile = Join-Path (Get-Location).Path $MetricsFile }
    $metricsAll = Read-PlMetrics $MetricsFile
}
if (-not $PSBoundParameters.ContainsKey('MinSeverity')) {
    if ($metricsAll) { $MinSeverity = 'Medium' } else { $MinSeverity = 'Info' }
}
$sevRank = @{ Info = 0; Medium = 1; High = 2 }
$minRank = $sevRank[$MinSeverity]

if (-not $SuppressFile) {
    $cand = Join-Path $script:PlOutDir 'perflint-suppress.tsv'
    if (Test-Path -LiteralPath $cand) { $SuppressFile = $cand }
}
$script:PlSuppress = Read-PlSuppress $SuppressFile

# список объектов
$targets = New-Object System.Collections.ArrayList
if ($All) {
    $idxPath = Join-Path $srcRoot 'index.tsv'
    if (-not (Test-Path -LiteralPath $idxPath)) { throw "Нет индекса дампа: $idxPath. Сначала Dump-AlSource.ps1." }
    $lines = [System.IO.File]::ReadAllLines($idxPath, [System.Text.Encoding]::UTF8)
    $rowsIdx = New-Object System.Collections.ArrayList
    for ($i = 1; $i -lt $lines.Length; $i++) {
        $c = $lines[$i] -split "`t"
        if ($c.Length -lt 5) { continue }
        $t0 = [int]$c[0]
        if ($Types -and $Types -notcontains $t0) { continue }
        [void]$rowsIdx.Add([pscustomobject]@{ Type = $t0; Id = [int]$c[2]; Name = $c[3]; Lines = [int]$c[4] })
    }
    $sorted = @($rowsIdx | Sort-Object -Property @{Expression='Lines';Descending=$true})
    if ($Top -gt 0 -and $sorted.Count -gt $Top) { $sorted = @($sorted | Select-Object -First $Top) }
    foreach ($r in $sorted) { [void]$targets.Add($r) }
}
else {
    $nm = ''
    $inf = Get-AlObjectInfo -ObjectType $ObjectType -ObjectId $ObjectId -SourceRoot $srcRoot
    if ($inf) { $nm = $inf.Name }
    [void]$targets.Add([pscustomobject]@{ Type = $ObjectType; Id = $ObjectId; Name = $nm; Lines = 0 })
}

if (-not $OutFile) {
    if ($All) { $OutFile = Join-Path $script:PlOutDir 'perflint-all.tsv' }
    else      { $OutFile = Join-Path $script:PlOutDir 'hints.tsv' }
}

$allFind  = New-Object System.Collections.ArrayList
$dropped  = 0
$errors   = New-Object System.Collections.ArrayList
$objDone  = 0
$objRank  = New-Object System.Collections.ArrayList

foreach ($tg in $targets) {
    $swObj = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $f = Invoke-PlObject -ObjType $tg.Type -ObjId $tg.Id -MetricsAll $metricsAll -SrcRoot $srcRoot -BaseDir $baseDir
        $objDone++
        $kept = 0
        foreach ($x in $f) {
            if ($sevRank[$x.Severity] -lt $minRank) { $dropped++; continue }
            [void]$allFind.Add($x)
            $kept++
        }
        $swObj.Stop()
        [void]$objRank.Add([pscustomobject]@{ Type = $tg.Type; Id = $tg.Id; Name = $tg.Name; Found = $kept; Sec = $swObj.Elapsed.TotalSeconds })
        if ($All -and -not $Quiet -and ($objDone % 25 -eq 0)) {
            Write-Host ('  {0} объектов, находок {1}...' -f $objDone, $allFind.Count) -ForegroundColor DarkGray
        }
    }
    catch {
        $swObj.Stop()
        [void]$errors.Add(('{0} {1}: {2}' -f (Get-AlTypeName $tg.Type), $tg.Id, $_.Exception.Message))
    }
}

# ------------------------------------------------------------------ TSV
$tab = [char]9
$outRows = New-Object System.Collections.Generic.List[string]
$outRows.Add((@('Строка', 'RuleID', 'Тип', 'ID', 'Функция', 'Элемент', 'Сообщение', 'Severity') -join $tab))
foreach ($x in @($allFind | Sort-Object @{Expression='TypeName'}, @{Expression='ObjectId'}, @{Expression='LineNo'})) {
    $outRows.Add((@($x.LineNo, $x.RuleId, $x.TypeName, $x.ObjectId, $x.Function, $x.Element, $x.Message, $x.Severity) -join $tab))
}
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutFile, (($outRows -join "`r`n") + "`r`n"), $utf8NoBom)

$swTotal.Stop()

# ------------------------------------------------------------------ сводка
$cntHigh = 0; $cntMed = 0; $cntInfo = 0
foreach ($x in $allFind) {
    switch ($x.Severity) {
        'High'   { $cntHigh++ }
        'Medium' { $cntMed++ }
        default  { $cntInfo++ }
    }
}
$mtxt = 'нет (статический прогон, все находки Info)'
if ($metricsAll) { $mtxt = $MetricsFile }
Write-Host ''
Write-Host ('Объектов:  {0}; находок {1} (High {2} / Medium {3} / Info {4}); подавлено {5}; отсечено порогом {6}' -f `
    $objDone, $allFind.Count, $cntHigh, $cntMed, $cntInfo, $script:PlSuppressed, $dropped)
Write-Host ('Порог:     {0}; метрики: {1}' -f $MinSeverity, $mtxt)
Write-Host ('Файл:      {0}' -f $OutFile)
Write-Host ('Время:     {0:N1} с' -f $swTotal.Elapsed.TotalSeconds)

if ($allFind.Count -gt 0 -and -not $Quiet) {
    Write-Host ''
    Write-Host '--- по правилам ---' -ForegroundColor Yellow
    foreach ($g in ($allFind | Group-Object RuleId | Sort-Object { [int]$_.Name })) {
        $nm2 = ''
        if ($script:PlRuleNames.ContainsKey($g.Name)) { $nm2 = $script:PlRuleNames[$g.Name] }
        $h = @($g.Group | Where-Object { $_.Severity -eq 'High' }).Count
        Write-Host ('  {0}  x{1,-5} High {2,-4} {3}' -f $g.Name, $g.Count, $h, $nm2)
    }
}

if (-not $All -and $allFind.Count -gt 0 -and -not $Quiet) {
    Write-Host ''
    Write-Host '--- топ находок ---' -ForegroundColor Yellow
    $top15 = @($allFind | Sort-Object @{Expression={$sevRank[$_.Severity]};Descending=$true},
        @{Expression='GainMs';Descending=$true}, @{Expression='LineNo'}) | Select-Object -First 15
    foreach ($x in $top15) {
        $t = $x.Message
        if ($t.Length -gt 100) { $t = $t.Substring(0, 98) + '..' }
        Write-Host ('  [{0,-6}] стр {1,5} r{2} {3}' -f $x.Severity, $x.LineNo, $x.RuleId, $t)
    }
}

if ($All -and -not $Quiet) {
    Write-Host ''
    Write-Host '--- топ объектов по находкам ---' -ForegroundColor Yellow
    foreach ($r in @($objRank | Sort-Object @{Expression='Found';Descending=$true} | Select-Object -First 10)) {
        Write-Host ('  {0,-9} {1,-9} x{2,-5} {3:N1} с  {4}' -f (Get-AlTypeName $r.Type), $r.Id, $r.Found, $r.Sec, $r.Name)
    }
}

if ($errors.Count -gt 0) {
    Write-Host ''
    Write-Host ('Не разобрано объектов: {0}' -f $errors.Count) -ForegroundColor Yellow
    foreach ($e in @($errors | Select-Object -First 5)) { Write-Host ('  ' + $e) -ForegroundColor DarkYellow }
}

exit 0
