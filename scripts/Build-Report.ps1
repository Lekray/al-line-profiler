<#
.SYNOPSIS
    Строит HTML-отчёт по объекту NAV: листинг кода с длительностью выполнения каждой строки.

.DESCRIPTION
    Читает дамп исходников, сделанный Dump-AlSource.ps1, и рендерит самодостаточный HTML:
    ни внешних файлов, ни интернета — страницу можно открыть где угодно.

    Без -MetricsFile отчёт показывает только листинг с платформенной нумерацией строк.
    С метриками добавляются колонки Self, Total, число вызовов и SQL, тепловая заливка
    по собственному времени и раздел «Топ строк». С -HintsFile добавляется колонка
    подсказок по оптимизации.

    ФОРМАТ МЕТРИК (lines.tsv, UTF-8, CRLF, разделитель — табуляция), колонки по имени:
        ObjectType ObjectId LineNo Hits TotalMs SelfMs MinMs MaxMs SqlCount SqlMs Stmts
    Лишние колонки игнорируются, отсутствующие считаются нулём.

    ФОРМАТ ПОДСКАЗОК (hints.tsv) — те же 7 колонок, что пишет Invoke-PerfLint.ps1:
        Строка RuleID Тип ID Функция Элемент Сообщение
    Необязательные дополнительные колонки: Severity, Confidence, GainMs.

.PARAMETER ObjectType
    Тип объекта: 1 Table, 3 Report, 5 Codeunit, 6 XMLport, 8 Page, 9 Query.

.PARAMETER ObjectId
    Номер объекта.

.PARAMETER MetricsFile
    lines.tsv с построчными метриками. Необязателен.

.PARAMETER HintsFile
    hints.tsv с подсказками по оптимизации. Необязателен.

.PARAMETER OutFile
    Куда писать отчёт. По умолчанию <корень репозитория>\out\<тип>_<id>.html

.PARAMETER TopCount
    Сколько строк показать в разделе «Топ строк». По умолчанию 25.

.PARAMETER Open
    Открыть готовый отчёт в браузере.

.EXAMPLE
    .\Build-Report.ps1 -ObjectType 5 -ObjectId 80
    Листинг кодюнита Sales-Post без таймингов.

.EXAMPLE
    .\Build-Report.ps1 -ObjectType 5 -ObjectId 80 -MetricsFile ..\out\lines.tsv -HintsFile ..\out\hints.tsv -Open
    Полный отчёт: время по строкам, подсказки, топ строк.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][int] $ObjectType,
    [Parameter(Mandatory)][int] $ObjectId,
    [string] $SourceRoot,
    [string] $MetricsFile,
    [string] $HintsFile,
    [string] $OutFile,
    [int]    $TopCount = 25,
    [switch] $Open
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Lib-AlListing.ps1')

# --- вспомогательное --------------------------------------------------------
function Read-Tsv {
    <#
    .SYNOPSIS
        Читает TSV с заголовком в список хэш-таблиц «колонка -> значение».
    #>
    param([string]$Path)
    $rows = New-Object System.Collections.Generic.List[object]
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    if ($lines.Length -lt 1) { return ,$rows }
    $head = $lines[0] -split "`t"
    for ($i = 1; $i -lt $lines.Length; $i++) {
        if (-not $lines[$i].Trim()) { continue }
        $c = $lines[$i] -split "`t"
        $h = @{}
        for ($k = 0; $k -lt $head.Length; $k++) {
            if ($k -lt $c.Length) { $h[$head[$k]] = $c[$k] } else { $h[$head[$k]] = '' }
        }
        $rows.Add($h)
    }
    return ,$rows
}

function Get-Num {
    param($Hash, [string]$Key)
    if (-not $Hash.ContainsKey($Key)) { return 0.0 }
    $v = $Hash[$Key]
    if ([string]::IsNullOrWhiteSpace($v)) { return 0.0 }
    $out = 0.0
    if ([double]::TryParse(($v -replace ',', '.'), [Globalization.NumberStyles]::Float,
                           [Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return $out }
    return 0.0
}

function Format-Ms {
    param([double]$V)
    if ($V -le 0) { return '' }
    if ($V -ge 100) { return ('{0:N0}' -f $V) }
    return ('{0:N1}' -f $V)
}

# --- данные -----------------------------------------------------------------
$info = Get-AlObjectInfo -ObjectType $ObjectType -ObjectId $ObjectId -SourceRoot $SourceRoot
if (-not $info) { throw "Объект $ObjectType/$ObjectId отсутствует в индексе дампа." }

$listing = @(Get-AlListing -ObjectType $ObjectType -ObjectId $ObjectId -SourceRoot $SourceRoot)
$funcs   = @(Get-AlFunctionMap -Listing $listing)

# --- метрики ----------------------------------------------------------------
$metrics    = @{}      # LineNo -> хэш метрик
$hasTimings = $false
if ($MetricsFile) {
    if (Test-Path -LiteralPath $MetricsFile) {
        foreach ($r in (Read-Tsv $MetricsFile)) {
            if ($r.ContainsKey('ObjectType') -and $r.ContainsKey('ObjectId')) {
                if ([int](Get-Num $r 'ObjectType') -ne $ObjectType) { continue }
                if ([int](Get-Num $r 'ObjectId')   -ne $ObjectId)   { continue }
            }
            $ln = [int](Get-Num $r 'LineNo')
            if ($ln -le 0) { continue }
            $metrics[$ln] = @{
                Hits     = [int](Get-Num $r 'Hits')
                TotalMs  = Get-Num $r 'TotalMs'
                SelfMs   = Get-Num $r 'SelfMs'
                SqlCount = [int](Get-Num $r 'SqlCount')
                SqlMs    = Get-Num $r 'SqlMs'
                Stmts    = [int](Get-Num $r 'Stmts')
            }
        }
        $hasTimings = ($metrics.Count -gt 0)
        if (-not $hasTimings) { Write-Warning "В $MetricsFile нет строк по объекту $ObjectType/$ObjectId." }
    } else {
        Write-Warning "Файл метрик не найден: $MetricsFile. Отчёт будет без таймингов."
    }
}

# --- подсказки --------------------------------------------------------------
$hints    = @{}        # LineNo -> список подсказок
$hintsAll = 0
if ($HintsFile) {
    if (Test-Path -LiteralPath $HintsFile) {
        foreach ($r in (Read-Tsv $HintsFile)) {
            if ($r.ContainsKey('ID') -and $r['ID'] -match '^\d+$' -and [int]$r['ID'] -ne $ObjectId) { continue }
            $ln = 0
            foreach ($k in @('Строка', 'LineNo', 'Line')) { if ($r.ContainsKey($k)) { $ln = [int](Get-Num $r $k); break } }
            if ($ln -le 0) { continue }
            $rule = ''; foreach ($k in @('RuleID', 'Rule', 'Код')) { if ($r.ContainsKey($k)) { $rule = $r[$k]; break } }
            $msg  = ''; foreach ($k in @('Сообщение', 'Message', 'Текст')) { if ($r.ContainsKey($k)) { $msg = $r[$k]; break } }
            $sev  = 'Medium'; foreach ($k in @('Severity', 'Серьёзность')) { if ($r.ContainsKey($k) -and $r[$k]) { $sev = $r[$k]; break } }
            if (-not $hints.ContainsKey($ln)) { $hints[$ln] = New-Object System.Collections.Generic.List[object] }
            $hints[$ln].Add([pscustomobject]@{ Rule = $rule; Message = $msg; Severity = $sev })
            $hintsAll++
        }
    } else {
        Write-Warning "Файл подсказок не найден: $HintsFile."
    }
}

if (-not $OutFile) {
    $outDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'out'
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $OutFile = Join-Path $outDir ("{0}_{1}.html" -f $ObjectType, $ObjectId)
}

# --- агрегаты ---------------------------------------------------------------
$sumSelf = 0.0; $sumSql = 0.0; $maxSelf = 0.0
foreach ($m in $metrics.Values) {
    $sumSelf += $m.SelfMs; $sumSql += $m.SqlMs
    if ($m.SelfMs -gt $maxSelf) { $maxSelf = $m.SelfMs }
}
# Вклад строки - это её Self, и БОЛЬШЕ НИЧЕГО. SqlMs не слагаемое, а часть Self:
# запрос выполняется внутри самого оператора, из Self он не вычитается (Lib-AlTrace.ps1,
# шапка: «чистое время C/AL - SelfNoSqlMs»). Сложение Self+Sql удваивало SQL и смещало
# ранжирование в пользу SQL-строк вплоть до двух раз.
$attrTotal = $sumSelf

# Подсказки уровня Info в счётчик не входят: линтер по миллиону строк выдаёт их
# пачками, и они утопят три-пять настоящих находок. В подсказке они остаются.
function Get-CountableHints {
    param($List)
    # Запятая обязательна: без неё результат из одного элемента разворачивается
    # в сам объект, а у PSCustomObject в Windows PowerShell 5.1 нет свойства .Count,
    # и счётчик подсказок молча оказывается пустым.
    if (-not $List) { return ,@() }
    return ,@($List | Where-Object { $_.Severity -notmatch 'Info|Инфо' })
}

# свёртка по функциям: собственное время, SQL и число подсказок внутри тела
$fnAgg   = @{}
$covered = @{}    # номер строки -> функция реально профилировалась
foreach ($f in $funcs) {
    $s = 0.0; $q = 0.0; $h = 0; $hits = 0; $any = $false
    for ($ln = $f.FirstLine; $ln -le $f.LastLine; $ln++) {
        if ($metrics.ContainsKey($ln)) {
            $any = $true
            $s += $metrics[$ln].SelfMs; $q += $metrics[$ln].SqlMs
            if ($metrics[$ln].Hits -gt $hits) { $hits = $metrics[$ln].Hits }
        }
        if ($hints.ContainsKey($ln)) { $h += (Get-CountableHints $hints[$ln]).Count }
    }
    $fnAgg[$f.HeaderLine] = @{ SelfMs = $s; SqlMs = $q; Hints = $h; MaxHits = $hits }
    # Отметка «не выполнялась» осмысленна только внутри функций, до которых
    # прогон вообще дошёл. В непокрытых функциях она означала бы не то.
    if ($any) {
        for ($ln = $f.HeaderLine; $ln -le $f.LastLine; $ln++) { $covered[$ln] = $true }
    }
}

function Get-HeatClass {
    <#
    .SYNOPSIS
        Класс тепловой заливки по собственному времени строки.
    .DESCRIPTION
        Градиент строится относительно САМОЙ дорогой строки — иначе на профиле,
        где время размазано по сотне строк, вся карта окажется одного цвета.
        Но верхние ступени требуют и заметной доли в прогоне: строка, которая
        просто «самая дорогая из дешёвых», кричать не должна.
    #>
    param([double]$Self)
    if ($Self -le 0 -or $maxSelf -le 0) { return '' }
    $ofMax   = $Self / $maxSelf
    $ofTotal = if ($attrTotal -gt 0) { $Self / $attrTotal } else { 0 }
    $level =
        if     ($ofMax -ge 0.80) { 5 }
        elseif ($ofMax -ge 0.50) { 4 }
        elseif ($ofMax -ge 0.25) { 3 }
        elseif ($ofMax -ge 0.10) { 2 }
        else                     { 1 }
    if ($ofTotal -lt 0.01 -and $level -gt 2) { $level = 2 }
    return " h$level"
}

# --- сборка HTML ------------------------------------------------------------
$sb = New-Object System.Text.StringBuilder
function Add-Line { param([string]$s) [void]$sb.AppendLine($s) }

$title = "{0} {1} {2}" -f $info.TypeName, $info.ObjectId, $info.Name
$cols  = if ($hasTimings) { 7 } else { 2 }

Add-Line '<!DOCTYPE html>'
Add-Line '<html lang="ru"><head><meta charset="utf-8">'
Add-Line ('<title>{0}</title>' -f (ConvertTo-HtmlText $title))
Add-Line '<style>'
Add-Line @'
:root{--paper:#F5F7FA;--surface:#fff;--surface2:#EDF0F5;--ink:#151920;--ink2:#39414F;
--muted:#5D6879;--rule:#D8DEE7;--accent:#2C4A7C;--hl:#FFF3C4;--num:#9AA5B4;
--h1:#FBEBC8;--h2:#F8D08A;--h3:#F0A055;--h4:#DA6A3C;--h5:#B33A2B;--hink:#2A1C13;--hink5:#FFF3EE;
--sev-hi:#8E2F3A;--sev-md:#B5721F;--sev-lo:#5D6879;}
@media (prefers-color-scheme:dark){:root{--paper:#11141A;--surface:#171B23;--surface2:#1E2430;
--ink:#E6EAF1;--ink2:#BAC3D1;--muted:#8F9AAB;--rule:#2A3140;--accent:#89ACE0;--hl:#3A3320;--num:#5C6779;
--h1:#4A3A1C;--h2:#6B4A1E;--h3:#8A5424;--h4:#A85434;--h5:#C24A38;--hink:#FCEFE6;--hink5:#FFF3EE;
--sev-hi:#E4909A;--sev-md:#E0B36A;--sev-lo:#8F9AAB;}}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);
font-family:"Segoe UI",system-ui,sans-serif;font-size:15px;line-height:1.5}
header{background:var(--surface);border-bottom:1px solid var(--rule);padding:14px 20px;position:sticky;top:0;z-index:5}
h1{margin:0 0 4px;font-size:19px;font-weight:600}
h2{font-size:15px;margin:0 0 8px;font-weight:600}
.meta{color:var(--muted);font-size:12.5px;display:flex;flex-wrap:wrap;gap:6px 18px}
.meta b{color:var(--ink2);font-weight:600}
.note{background:var(--hl);color:var(--ink);padding:7px 20px;font-size:13px;border-bottom:1px solid var(--rule)}
.wrap{display:grid;grid-template-columns:270px minmax(0,1fr);align-items:start}
@media (max-width:980px){.wrap{grid-template-columns:1fr}nav{display:none}}
nav{position:sticky;top:74px;max-height:calc(100vh - 74px);overflow:auto;
border-right:1px solid var(--rule);padding:12px 10px 40px;background:var(--surface)}
nav .cap{font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);font-weight:700;padding:0 8px 8px}
nav a{display:flex;justify-content:space-between;gap:8px;text-decoration:none;color:var(--ink2);
padding:3px 8px;border-radius:4px;font-size:13px;font-family:Consolas,monospace}
nav a:hover{background:var(--surface2);color:var(--accent)}
nav a .n{color:var(--num);font-size:11.5px;flex:none;font-variant-numeric:tabular-nums}
main{padding:0 0 60px;min-width:0}
.tools{padding:10px 20px;border-bottom:1px solid var(--rule);display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.tools input{flex:1;min-width:220px;max-width:420px;padding:6px 10px;border:1px solid var(--rule);
border-radius:5px;background:var(--surface);color:var(--ink);font-size:14px}
.tools button{padding:6px 12px;border:1px solid var(--rule);border-radius:5px;background:var(--surface);
color:var(--ink2);cursor:pointer;font-size:13px}
.tools button:hover{border-color:var(--accent);color:var(--accent)}
.top{padding:14px 20px;border-bottom:1px solid var(--rule);background:var(--surface)}
table{border-collapse:collapse;width:100%}
table.src{font-family:Consolas,"Cascadia Mono",monospace;font-size:13px}
table.src td{padding:0 8px;vertical-align:top;white-space:pre;border:0}
table.src td.code{width:99%}
td.ln{text-align:right;color:var(--num);width:1%;user-select:none;border-right:1px solid var(--rule);
font-variant-numeric:tabular-nums;padding-right:10px}
td.m{text-align:right;width:1%;font-variant-numeric:tabular-nums;white-space:nowrap;color:var(--ink2)}
td.m.self{border-left:1px solid var(--rule)}
td.self.h1{background:var(--h1);color:var(--hink)}
td.self.h2{background:var(--h2);color:var(--hink)}
td.self.h3{background:var(--h3);color:var(--hink)}
td.self.h4{background:var(--h4);color:var(--hink)}
td.self.h5{background:var(--h5);color:var(--hink5)}
tr.fn td{background:var(--surface2);font-weight:700;padding-top:14px;padding-bottom:4px}
tr.cm td.code{color:var(--muted)}
tr:target td.code{background:var(--hl)}
tr.hit td.code{background:var(--hl)}
tr.hide{display:none}
a.anchor{color:inherit;text-decoration:none}
.chip{display:inline-block;font-family:"Segoe UI",sans-serif;font-size:11px;font-weight:700;
padding:0 7px;border-radius:9px;cursor:help;color:#fff}
.chip.hi{background:var(--sev-hi)}
.chip.md{background:var(--sev-md)}
.chip.lo{background:var(--sev-lo)}
table.top td,table.top th{padding:4px 10px;font-size:13px;border-bottom:1px solid var(--rule);text-align:right;
font-variant-numeric:tabular-nums}
table.top th{font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);text-align:right}
table.top td.c,table.top th.c{text-align:left;font-family:Consolas,monospace;white-space:pre;
overflow:hidden;text-overflow:ellipsis;max-width:0;width:99%}
table.top a{color:var(--accent);text-decoration:none;font-variant-numeric:tabular-nums}
'@
Add-Line '</style></head><body>'

# --- шапка ------------------------------------------------------------------
Add-Line '<header>'
Add-Line ('<h1>{0}</h1>' -f (ConvertTo-HtmlText $title))
Add-Line '<div class="meta">'
Add-Line ('<span>строк: <b>{0}</b></span>' -f $info.Lines)
Add-Line ('<span>функций: <b>{0}</b></span>' -f $funcs.Count)
$compiledText = if ($info.Compiled) { 'да' } else { 'НЕТ — нумерация строк может не соответствовать коду' }
Add-Line ('<span>скомпилирован: <b>{0}</b></span>' -f $compiledText)
Add-Line ('<span>изменён: <b>{0} {1}</b></span>' -f $info.Date, $info.Time)
if ($hasTimings) {
    Add-Line ('<span>строк с данными: <b>{0}</b></span>' -f $metrics.Count)
    Add-Line ('<span>Self: <b>{0} мс</b></span>' -f (Format-Ms $sumSelf))
    Add-Line ('<span>в т.ч. SQL: <b>{0} мс</b></span>' -f (Format-Ms $sumSql))
}
if ($hintsAll -gt 0) { Add-Line ('<span>подсказок: <b>{0}</b></span>' -f $hintsAll) }
Add-Line ('<span>нумерация: <b>платформенная</b></span>')
Add-Line '</div></header>'
if (-not $hasTimings) {
    Add-Line '<div class="note">Тайминги не собраны — показан только листинг. Колонки времени появятся после прогона трассировки.</div>'
}

Add-Line '<div class="wrap">'

# --- список функций ---------------------------------------------------------
Add-Line '<nav><div class="cap">Функции</div>'
foreach ($f in $funcs) {
    $suffix = ('<span class="n">{0}</span>' -f $f.HeaderLine)
    if ($hasTimings -and $fnAgg[$f.HeaderLine].SelfMs -gt 0) {
        $suffix = ('<span class="n">{0} мс</span>' -f (Format-Ms $fnAgg[$f.HeaderLine].SelfMs))
    }
    Add-Line ('<a href="#L{0}"><span>{1}</span>{2}</a>' -f $f.HeaderLine, (ConvertTo-HtmlText $f.Name), $suffix)
}
Add-Line '</nav>'

Add-Line '<main>'

# --- топ строк --------------------------------------------------------------
if ($hasTimings) {
    $top = @()
    foreach ($ln in $metrics.Keys) {
        $m = $metrics[$ln]
        $a = $m.SelfMs
        if ($a -le 0) { continue }
        $top += [pscustomobject]@{ LineNo = $ln; Attr = $a; M = $m }
    }
    $top = @($top | Sort-Object -Property Attr -Descending | Select-Object -First $TopCount)
    if ($top.Count -gt 0) {
        Add-Line '<div class="top">'
        Add-Line ('<h2>Топ строк по собственному времени, первые {0}</h2>' -f $top.Count)
        Add-Line '<table class="top"><thead><tr><th>Строка</th><th>Self, мс</th><th>в т.ч. SQL, мс</th><th>Вызовов</th><th>% вклада</th><th class="c">Код</th></tr></thead><tbody>'
        foreach ($t in $top) {
            $txt = ($listing[$t.LineNo - 1].Text).Trim()
            $pct = if ($attrTotal -gt 0) { 100.0 * $t.Attr / $attrTotal } else { 0 }
            $ha = ''
            if ($t.M.Stmts -gt 1) { $ha = ' title="операторов в строке: {0}; показана сумма попаданий по ним"' -f $t.M.Stmts }
            Add-Line ('<tr><td><a href="#L{0}">{0}</a></td><td>{1}</td><td>{2}</td><td{3}>{4:N0}</td><td>{5:N1}</td><td class="c">{6}</td></tr>' -f `
                $t.LineNo, (Format-Ms $t.M.SelfMs), (Format-Ms $t.M.SqlMs), $ha, $t.M.Hits, $pct, (ConvertTo-HtmlText $txt))
        }
        Add-Line '</tbody></table></div>'
    }
}

# --- листинг ----------------------------------------------------------------
Add-Line '<div class="tools">'
Add-Line '<input id="q" type="search" placeholder="Поиск по коду — введите текст и нажмите Enter">'
Add-Line '<button onclick="clr()">Сбросить</button>'
Add-Line '<span id="cnt" style="color:var(--muted);font-size:13px"></span>'
Add-Line '</div>'
Add-Line '<table class="src"><tbody>'

foreach ($l in $listing) {
    $cls = switch ($l.Kind) { 'Function' { ' class="fn"' } 'Comment' { ' class="cm"' } default { '' } }
    $cells = ('<td class="ln"><a class="anchor" href="#L{0}">{0}</a></td>' -f $l.LineNo)

    if ($hasTimings) {
        $self = ''; $tot = ''; $hits = ''; $sql = ''; $heat = ''; $hitsAttr = ''
        if ($l.Kind -eq 'Function') {
            $a = $fnAgg[$l.LineNo]
            if ($a) {
                $self = Format-Ms $a.SelfMs
                $sql  = Format-Ms $a.SqlMs
                $hits = if ($a.MaxHits -gt 0) { '{0:N0}' -f $a.MaxHits } else { '' }
            }
        }
        elseif ($metrics.ContainsKey($l.LineNo)) {
            $m = $metrics[$l.LineNo]
            $self = Format-Ms $m.SelfMs
            $tot  = Format-Ms $m.TotalMs
            $hits = if ($m.Hits -gt 0) { '{0:N0}' -f $m.Hits } else { '' }
            # Событие приходит на КАЖДЫЙ оператор, а на строке их может быть несколько
            # («IF ... THEN ...;»). Тогда в колонке лежит сумма попаданий по операторам
            # строки, а не число её выполнений: делить нельзя — условный оператор
            # срабатывает не всегда. Разница объясняется подсказкой при наведении.
            if ($m.Stmts -gt 1) {
                $hitsAttr = ' title="операторов в строке: {0}; показана сумма попаданий по ним"' -f $m.Stmts
            }
            $sql  = Format-Ms $m.SqlMs
            $heat = Get-HeatClass $m.SelfMs
        }
        elseif ($l.Kind -eq 'Code' -and $covered.ContainsKey($l.LineNo)) {
            # функция профилировалась, а эта её строка не выполнилась ни разу
            $self = '·'
        }
        $cells += ('<td class="m self{0}">{1}</td><td class="m">{2}</td><td class="m"{3}>{4}</td><td class="m">{5}</td>' -f `
            $heat, $self, $tot, $hitsAttr, $hits, $sql)

        # колонка подсказок
        $hc = ''
        if ($l.Kind -eq 'Function') {
            $a = $fnAgg[$l.LineNo]
            if ($a -and $a.Hints -gt 0) { $hc = ('<span class="chip lo" title="подсказок внутри функции">{0}</span>' -f $a.Hints) }
        }
        elseif ($hints.ContainsKey($l.LineNo)) {
            $hl  = $hints[$l.LineNo]
            $cnt = Get-CountableHints $hl
            if ($cnt.Count -gt 0) {
                $sev = 'lo'
                foreach ($x in $cnt) {
                    if ($x.Severity -match 'Critical|High|Крит|Выс') { $sev = 'hi'; break }
                    if ($x.Severity -match 'Medium|Сред') { $sev = 'md' }
                }
                # в подсказке показываем всё, включая информационные
                $tip = (($hl | ForEach-Object { ('{0}: {1}' -f $_.Rule, $_.Message) }) -join ' | ')
                $hc = ('<span class="chip {0}" title="{1}">{2}</span>' -f $sev, (ConvertTo-HtmlText $tip), $cnt.Count)
            }
        }
        $cells += ('<td class="m">{0}</td>' -f $hc)
    }

    $cells += ('<td class="code">{0}</td>' -f (ConvertTo-HtmlText $l.Text))
    Add-Line ('<tr id="L{0}"{1}>{2}</tr>' -f $l.LineNo, $cls, $cells)
}

Add-Line '</tbody></table></main></div>'

# --- поиск ------------------------------------------------------------------
Add-Line ('<script>var CODECOL = {0};' -f ($cols - 1))
Add-Line @'
var rows = document.querySelectorAll('table.src tr');
function clr(){ document.getElementById('q').value='';
  for (var i=0;i<rows.length;i++){ rows[i].classList.remove('hide','hit'); }
  document.getElementById('cnt').textContent=''; }
document.getElementById('q').addEventListener('keydown', function(e){
  if (e.key !== 'Enter') return;
  var q = this.value.toLowerCase();
  if (!q) { clr(); return; }
  var n = 0;
  for (var i=0;i<rows.length;i++){
    var cell = rows[i].cells[rows[i].cells.length-1];
    var ok = cell.textContent.toLowerCase().indexOf(q) >= 0;
    rows[i].classList.toggle('hit', ok);
    rows[i].classList.toggle('hide', !ok && !rows[i].classList.contains('fn'));
    if (ok) n++;
  }
  document.getElementById('cnt').textContent = 'найдено строк: ' + n;
});
</script>
'@
Add-Line '</body></html>'

# --- запись -----------------------------------------------------------------
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), $utf8NoBom)

$size = (Get-Item -LiteralPath $OutFile).Length
Write-Host ''
Write-Host ("Объект:  {0}" -f $title)
Write-Host ("Строк:   {0}" -f $info.Lines)
Write-Host ("Функций: {0}" -f $funcs.Count)
if ($hasTimings) {
    Write-Host ("Метрики: {0} строк, Self {1} мс, SQL {2} мс" -f $metrics.Count, (Format-Ms $sumSelf), (Format-Ms $sumSql))
} else {
    Write-Host 'Метрики: не собраны' -ForegroundColor Yellow
}
if ($hintsAll -gt 0) { Write-Host ("Подсказок: {0}" -f $hintsAll) }
Write-Host ("Отчёт:   {0} ({1:N0} КБ)" -f $OutFile, ($size / 1KB))
if ($Open) { Start-Process $OutFile }
