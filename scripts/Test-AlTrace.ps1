#Requires -Version 5.1
<#
.SYNOPSIS
    Обкатка ядра разбора трассировки на потоках событий с заранее известным ответом.

.DESCRIPTION
    Отказы Lib-AlTrace.ps1 выглядят как успех: разбор не падает, предупреждений нет,
    просто числа выходят неправильные. Поймать такое живой трассировкой нельзя - там
    не с чем сверять. Поэтому здесь короткие искусственные потоки событий, для каждого
    из которых верный ответ посчитан руками и записан в проверке.

    Что закрывает (находки аудита 14.08.2026):
      №1  осиротевший SQL-Start отравлял список открытых операций, и учёт SQL-времени
          сессии выключался навсегда; поздний Stop той же операции спаривался с
          огрызком и писал фантомные секунды;
      №2  ошибка C/AL не снимала рамку, и строка вызова забирала себе весь остаток
          трассы;
      №4  вклад строки считался как Self + SQL, хотя SQL - часть Self, а не слагаемое.

    Судит сама: первая строка - "пройдено N из M", код возврата ненулевой при отказе.

.EXAMPLE
    powershell -File scripts\Test-AlTrace.ps1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Lib-AlTrace.ps1')

$MsTicks = 10000L

$script:Seq    = 0
$script:Passed = 0
$script:Total  = 0
$script:Fails  = New-Object System.Collections.Generic.List[string]

function New-Ev {
    <#
    .SYNOPSIS
        Событие трассировки в том виде, в каком его отдаёт Import-AlTraceEvents.
    #>
    param(
        [Parameter(Mandatory)][string] $Kind,
        [Parameter(Mandatory)][double] $Ms,
        [string] $Fn   = 'A',
        [int]    $Line = 0,
        [string] $Op   = '',
        [int]    $Oi   = 110200
    )
    $script:Seq++
    return [pscustomobject]@{
        Seq          = $script:Seq
        EventId      = 0
        Role         = $Kind
        Kind         = $Kind
        OpName       = $Op
        Ticks        = [int64]($Ms * $MsTicks)
        SessionId    = '1'
        ObjectType   = 5
        ObjectId     = $Oi
        FunctionName = $Fn
        LineNumber   = $Line
        Statement    = ('оператор строки {0}' -f $Line)
        RecordId     = [int64]$script:Seq
        ThreadId     = 1
        Sql          = ''
    }
}

function Test-Value {
    <#
    .SYNOPSIS
        Сверка одного числа с ожидаемым; дробные сравниваются с допуском.
    #>
    param([string] $Name, $Expected, $Actual, [double] $Tolerance = 0.0005)
    $script:Total++
    $ok = $false
    if ($Expected -is [double] -or $Expected -is [decimal] -or $Expected -is [single]) {
        $a = 0.0
        if ($null -ne $Actual) { $a = [double]$Actual }
        $ok = ([math]::Abs($a - [double]$Expected) -le $Tolerance)
    }
    else { $ok = ($Actual -eq $Expected) }

    if ($ok) { $script:Passed++ }
    else {
        $shown = if ($null -eq $Actual) { '<нет>' } else { $Actual }
        [void]$script:Fails.Add(('FAILED {0}: ожидалось {1}, получено {2}' -f $Name, $Expected, $shown))
    }
}

function Get-Row {
    <#
    .SYNOPSIS
        Строка результата по номеру; $null, если строки в результате нет вовсе.
    #>
    param($Rows, [int] $LineNo)
    $m = @($Rows | Where-Object { $_.LineNo -eq $LineNo })
    if ($m.Count -eq 0) { return $null }
    return $m[0]
}

# Смещение 0: в проверках номер строки листинга равен номеру из события.
function Measure-Case { param($Events) return @(Measure-AlLines -Events $Events -LineOffset 0) }

function New-TestDump {
    <#
    .SYNOPSIS
        Минимальный дамп исходников: заголовок функции и два оператора.
    #>
    param([string] $Root, [int] $Ot = 5, [int] $Oi = 110200)
    New-Item -ItemType Directory -Path $Root -Force | Out-Null
    $u = New-Object System.Text.UTF8Encoding($false)
    $t = [char]9
    [System.IO.File]::WriteAllText((Join-Path $Root ('{0}_{1}.al' -f $Ot, $Oi)),
        "MyFunc()`r`n  Tick;`r`n  Done := Done + 1;`r`n", $u)
    [System.IO.File]::WriteAllText((Join-Path $Root 'index.tsv'),
        (('ObjectType','TypeName','ObjectId','Name','Lines','Bytes','Compiled','Date','Time','VersionList','Hash') -join $t) + "`r`n" +
        ((([string]$Ot),'Codeunit',([string]$Oi),'Проверка','3','60','1','18.08.26','12:00:00','TEST','0') -join $t) + "`r`n", $u)
}

function New-TestEvents {
    <#
    .SYNOPSIS
        events.tsv в том виде, в каком его пишет сборщик трассировки.
    #>
    param([string] $Path, [object[]] $Rows)
    $t = [char]9
    $out = New-Object System.Collections.Generic.List[string]
    [void]$out.Add((('EventId','EventName','TimeCreatedTicks','SessionId','ObjectType','ObjectId',
                     'FunctionName','LineNumber','Statement','RecordId','ThreadId','Raw') -join $t))
    foreach ($r in $Rows) {
        [void]$out.Add(( '0', $r.Name, ([string][int64]($r.Ms * 10000)), '1',
                         ([string]$r.Ot), ([string]$r.Oi), $r.Fn, $r.Line, $r.Stmt,
                         ([string]$r.Rec), '1', '' ) -join $t)
    }
    [System.IO.File]::WriteAllText($Path, (($out -join "`r`n") + "`r`n"),
                                   (New-Object System.Text.UTF8Encoding($false)))
}

function Invoke-Rebuild {
    <#
    .SYNOPSIS
        Прогон Rebuild-Report.ps1 целиком; возвращает код возврата (-1, если упал).
    #>
    param([string] $RunDir, [string] $Src, [int] $Ot = 5, [int] $Oi = 110200)
    try {
        $global:LASTEXITCODE = 0
        & (Join-Path $PSScriptRoot 'Rebuild-Report.ps1') -RunDir $RunDir -ObjectType $Ot `
            -ObjectId $Oi -SourceRoot $Src -NoHints *>$null
        if ($null -eq $LASTEXITCODE) { return 0 }
        return [int]$LASTEXITCODE
    }
    catch { return -1 }
}

function Get-MetricLines {
    <#
    .SYNOPSIS
        Номера строк из lines.tsv прогона.
    #>
    param([string] $RunDir)
    $p = Join-Path $RunDir 'lines.tsv'
    if (-not (Test-Path -LiteralPath $p)) { return @() }
    $l = [System.IO.File]::ReadAllLines($p, [System.Text.Encoding]::UTF8)
    if ($l.Length -lt 2) { return @() }
    $h = $l[0] -split "`t"
    $ix = [array]::IndexOf($h, 'LineNo')
    if ($ix -lt 0) { return @() }
    $res = @()
    for ($i = 1; $i -lt $l.Length; $i++) {
        if (-not $l[$i].Trim()) { continue }
        $res += [int]($l[$i] -split "`t")[$ix]
    }
    return ($res | Sort-Object -Unique)
}

function New-TestRun {
    <#
    .SYNOPSIS
        Пустой каталог прогона рядом с дампом; возвращает оба пути.
    #>
    $root = Join-Path ([System.IO.Path]::GetTempPath()) ('lp-run-' + [guid]::NewGuid().ToString('N'))
    $src  = Join-Path $root 'src'
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    New-TestDump -Root $src
    return [pscustomobject]@{ Dir = $root; Src = $src }
}

# ---------------------------------------------------------------------------
# 1. Контроль: без SQL и без ошибок числа должны быть очевидными
# ---------------------------------------------------------------------------
$script:Seq = 0
$r = Measure-Case @(
    (New-Ev -Kind 'Start' -Ms 0),
    (New-Ev -Kind 'Stmt'  -Ms 1  -Line 10),
    (New-Ev -Kind 'Stmt'  -Ms 4  -Line 11),
    (New-Ev -Kind 'Stop'  -Ms 10)
)
Test-Value 'контроль: Total строки 10' 3.0  (Get-Row $r 10).TotalMs
Test-Value 'контроль: Self строки 10'  3.0  (Get-Row $r 10).SelfMs
Test-Value 'контроль: Total строки 11' 6.0  (Get-Row $r 11).TotalMs
Test-Value 'контроль: попаданий'       1L   (Get-Row $r 10).Hits

# ---------------------------------------------------------------------------
# 2. Контроль: парный SQL ложится на строку, внутри которой начался
# ---------------------------------------------------------------------------
$script:Seq = 0
$r = Measure-Case @(
    (New-Ev -Kind 'Start'     -Ms 0),
    (New-Ev -Kind 'Stmt'      -Ms 1  -Line 10),
    (New-Ev -Kind 'PairStart' -Ms 2  -Op 'SELECT'),
    (New-Ev -Kind 'PairStop'  -Ms 12 -Op 'SELECT'),
    (New-Ev -Kind 'Stmt'      -Ms 20 -Line 11),
    (New-Ev -Kind 'Stop'      -Ms 20)
)
Test-Value 'контроль SQL: SqlMs строки 10'    10.0 (Get-Row $r 10).SqlMs
Test-Value 'контроль SQL: запросов строки 10' 1L   (Get-Row $r 10).SqlCount
Test-Value 'контроль SQL: Total строки 10'    19.0 (Get-Row $r 10).TotalMs

# ---------------------------------------------------------------------------
# 3. Находка №1: огрызок SQL-Start не должен глушить учёт до конца сессии
# ---------------------------------------------------------------------------
$script:Seq = 0
$r = Measure-Case @(
    (New-Ev -Kind 'Start'     -Ms 0),
    (New-Ev -Kind 'Stmt'      -Ms 1  -Line 10),
    (New-Ev -Kind 'PairStart' -Ms 2  -Op 'SELECT'),   # пары не будет: событие потеряно
    (New-Ev -Kind 'Stmt'      -Ms 10 -Line 11),
    (New-Ev -Kind 'PairStart' -Ms 11 -Op 'SELECT'),
    (New-Ev -Kind 'PairStop'  -Ms 21 -Op 'SELECT'),
    (New-Ev -Kind 'Stop'      -Ms 30)
)
$s = $script:AlMeasureStats
Test-Value 'огрызок: SQL следующей строки учтён' 10.0 (Get-Row $r 11).SqlMs
Test-Value 'огрызок: запросов на строке 11'      1L   (Get-Row $r 11).SqlCount
Test-Value 'огрызок: на строке 10 времени нет'   0.0  (Get-Row $r 10).SqlMs
Test-Value 'огрызок: снят и посчитан'            1    $s.SqlDropped

# ---------------------------------------------------------------------------
# 4. Находка №1, обратная сторона: поздний Stop не спаривается с огрызком
# ---------------------------------------------------------------------------
$script:Seq = 0
$r = Measure-Case @(
    (New-Ev -Kind 'Start'     -Ms 0),
    (New-Ev -Kind 'Stmt'      -Ms 1   -Line 10),
    (New-Ev -Kind 'PairStart' -Ms 2   -Op 'SELECT'),  # пары не будет
    (New-Ev -Kind 'Stmt'      -Ms 10  -Line 11),
    (New-Ev -Kind 'PairStop'  -Ms 510 -Op 'SELECT'),  # чужой Stop полсекунды спустя
    (New-Ev -Kind 'Stop'      -Ms 520)
)
$s = $script:AlMeasureStats
Test-Value 'фантом: строка 10 без SQL-времени' 0.0 (Get-Row $r 10).SqlMs
Test-Value 'фантом: строка 11 без SQL-времени' 0.0 (Get-Row $r 11).SqlMs
Test-Value 'фантом: Stop признан беспарным'    1   $s.SqlUnmatched

# ---------------------------------------------------------------------------
# 5. Находка №2: ошибка C/AL снимает рамку, а не оставляет её до конца разбора
# ---------------------------------------------------------------------------
$script:Seq = 0
$r = Measure-Case @(
    (New-Ev -Kind 'Start' -Ms 0    -Fn 'A'),
    (New-Ev -Kind 'Stmt'  -Ms 1    -Fn 'A' -Line 10),   # строка вызова
    (New-Ev -Kind 'Start' -Ms 2    -Fn 'B'),
    (New-Ev -Kind 'Stmt'  -Ms 2    -Fn 'B' -Line 20),
    (New-Ev -Kind 'Error' -Ms 4    -Fn 'B'),
    (New-Ev -Kind 'Error' -Ms 4    -Fn 'A'),
    (New-Ev -Kind 'Start' -Ms 1000 -Fn 'C'),            # к ошибке отношения не имеет
    (New-Ev -Kind 'Stmt'  -Ms 1000 -Fn 'C' -Line 30),
    (New-Ev -Kind 'Stop'  -Ms 1002 -Fn 'C')
)
$s = $script:AlMeasureStats
Test-Value 'ошибка: Total строки вызова'     3.0 (Get-Row $r 10).TotalMs
Test-Value 'ошибка: Self строки вызова'      1.0 (Get-Row $r 10).SelfMs
Test-Value 'ошибка: Total строки в упавшей'  2.0 (Get-Row $r 20).TotalMs
Test-Value 'ошибка: рамок свёрнуто'          2   $s.ErrorPops
Test-Value 'ошибка: чужая рамка не задета'   2.0 (Get-Row $r 30).TotalMs

# ---------------------------------------------------------------------------
# 6. Перехваченная ошибка: вызывающий продолжает, и числа не портятся
# ---------------------------------------------------------------------------
$script:Seq = 0
$r = Measure-Case @(
    (New-Ev -Kind 'Start' -Ms 0  -Fn 'A'),
    (New-Ev -Kind 'Stmt'  -Ms 1  -Fn 'A' -Line 10),
    (New-Ev -Kind 'Start' -Ms 2  -Fn 'B'),
    (New-Ev -Kind 'Stmt'  -Ms 2  -Fn 'B' -Line 20),
    (New-Ev -Kind 'Error' -Ms 4  -Fn 'B'),
    (New-Ev -Kind 'Stmt'  -Ms 5  -Fn 'A' -Line 11),     # ошибку перехватили
    (New-Ev -Kind 'Stop'  -Ms 10 -Fn 'A')
)
$s = $script:AlMeasureStats
Test-Value 'перехват: Total строки вызова' 4.0 (Get-Row $r 10).TotalMs
Test-Value 'перехват: Self строки вызова'  2.0 (Get-Row $r 10).SelfMs
Test-Value 'перехват: Total следующей'     5.0 (Get-Row $r 11).TotalMs
Test-Value 'перехват: без грубой свёртки'  0   $s.ForcedPops

# ---------------------------------------------------------------------------
# 7. Находка №4, сквозняком: вклад строки в отчёте не удваивает SQL
# ---------------------------------------------------------------------------
# Числа подобраны так, что удвоение меняет ПОРЯДОК строк в топе, а это видно
# независимо от того, как система форматирует дробные.
#   строка 2: Self 100, SQL 90  -> с ошибкой вклад 190, верно 100
#   строка 3: Self 120, SQL  0  -> с ошибкой вклад 120, верно 120
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('lp-test-' + [guid]::NewGuid().ToString('N'))
$src = Join-Path $tmp 'src'
New-Item -ItemType Directory -Path $src -Force | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)
$tab  = [char]9

[System.IO.File]::WriteAllText((Join-Path $src '5_110200.al'),
    "MyFunc()`r`n  ПервыйОператор;`r`n  ВторойОператор;`r`n", $utf8)
[System.IO.File]::WriteAllText((Join-Path $src 'index.tsv'),
    (('ObjectType','TypeName','ObjectId','Name','Lines','Bytes','Compiled','Date','Time','VersionList','Hash') -join $tab) + "`r`n" +
    (('5','Codeunit','110200','Проверка','3','60','1','18.08.26','12:00:00','TEST','0') -join $tab) + "`r`n", $utf8)

$metrics = Join-Path $tmp 'lines.tsv'
[System.IO.File]::WriteAllText($metrics,
    (('ObjectType','ObjectId','LineNo','Hits','TotalMs','SelfMs','MinMs','MaxMs','SqlCount','SqlMs','Stmts') -join $tab) + "`r`n" +
    (('5','110200','2','10','200','100','1','20','5','90','1') -join $tab) + "`r`n" +
    (('5','110200','3','10','120','120','1','20','0','0','1')  -join $tab) + "`r`n", $utf8)

$html = Join-Path $tmp 'report.html'
& (Join-Path $PSScriptRoot 'Build-Report.ps1') -ObjectType 5 -ObjectId 110200 `
    -SourceRoot $src -MetricsFile $metrics -OutFile $html *>$null

if (-not (Test-Path -LiteralPath $html)) {
    $script:Total++
    [void]$script:Fails.Add('FAILED вклад: отчёт не собрался')
}
else {
    $text = [System.IO.File]::ReadAllText($html, [System.Text.Encoding]::UTF8)
    $topIx = $text.IndexOf('Топ строк')
    $first = -1
    if ($topIx -ge 0) {
        $m = [regex]::Match($text.Substring($topIx), '<tr><td><a href="#L(\d+)"')
        if ($m.Success) { $first = [int]$m.Groups[1].Value }
    }
    Test-Value 'вклад: первой идёт строка с большим Self' 3 $first
    Test-Value 'вклад: заголовок больше не обещает сумму' $false $text.Contains('Self + SQL')
}
Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 8. Находка №18: облегчённый сбор - не сбой. Событий уровня оператора в нём нет
#    по построению, и объявлять это красным «нет операторов объекта» нельзя.
# ---------------------------------------------------------------------------
$r8 = New-TestRun
New-TestEvents (Join-Path $r8.Dir 'events.tsv') @(
    @{ Name = 'SlowSql'; Ms = 1; Ot = 5; Oi = 110200; Fn = ''; Line = ''; Stmt = ''; Rec = 1 },
    @{ Name = 'SlowSql'; Ms = 2; Ot = 5; Oi = 110200; Fn = ''; Line = ''; Stmt = ''; Rec = 2 },
    @{ Name = 'SlowSql'; Ms = 3; Ot = 5; Oi = 110200; Fn = ''; Line = ''; Stmt = ''; Rec = 3 }
)
Test-Value 'облегчённый сбор: код возврата' 0 (Invoke-Rebuild -RunDir $r8.Dir -Src $r8.Src)
Remove-Item -LiteralPath $r8.Dir -Recurse -Force -ErrorAction SilentlyContinue

# Здоровый прогон: та же обвязка, всё на месте - контроль, чтобы девятка ниже
# не оказалась следствием самой обвязки.
$okRows = @(
    @{ Name = 'ALFunctionStart'; Ms = 1;  Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '';  Stmt = '';                  Rec = 1 },
    @{ Name = 'ALStatement';     Ms = 1;  Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '1'; Stmt = 'Tick;';             Rec = 2 },
    @{ Name = 'ALStatement';     Ms = 4;  Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '2'; Stmt = 'Done := Done + 1;'; Rec = 3 },
    @{ Name = 'ALFunctionStop';  Ms = 10; Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '';  Stmt = '';                  Rec = 4 }
)
$r9 = New-TestRun
New-TestEvents (Join-Path $r9.Dir 'events.tsv') $okRows
Test-Value 'здоровый прогон: код возврата' 0 (Invoke-Rebuild -RunDir $r9.Dir -Src $r9.Src)
Test-Value 'здоровый прогон: строки метрик' '2,3' ((Get-MetricLines $r9.Dir) -join ',')
Remove-Item -LiteralPath $r9.Dir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 9. В трассировке есть операторы, но чужого объекта - это другой сценарий в клиенте
# ---------------------------------------------------------------------------
$r10 = New-TestRun
New-TestEvents (Join-Path $r10.Dir 'events.tsv') (@($okRows | ForEach-Object {
    $c = $_.Clone(); $c.Oi = 110201; $c }))
Test-Value 'другой сценарий: код возврата' 8 (Invoke-Rebuild -RunDir $r10.Dir -Src $r10.Src)
Remove-Item -LiteralPath $r10.Dir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 10. Находки №21/25: гейт валидности прогона. Разрыв нумерации записей канала -
#     это потерянные события; отчёт строится, но верить его числам нельзя.
# ---------------------------------------------------------------------------
$r11 = New-TestRun
$gapRows = @($okRows | ForEach-Object { $c = $_.Clone(); $c })
$gapRows[3].Rec = 50          # пропуск записей 4..49
New-TestEvents (Join-Path $r11.Dir 'events.tsv') $gapRows
Test-Value 'гейт: разрыв нумерации даёт код 9' 9 (Invoke-Rebuild -RunDir $r11.Dir -Src $r11.Src)
Test-Value 'гейт: отчёт всё равно построен' $true (Test-Path -LiteralPath (Join-Path $r11.Dir '5_110200.html'))
Remove-Item -LiteralPath $r11.Dir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 11. Находки №8/20/24: провал калибровки включает резервный маппинг по тексту.
#     Оба оператора помечены одним номером строки, поэтому ни одно смещение из
#     -1/0/1 не даёт 80 % совпадений. По смещению оба легли бы на строку 2;
#     по тексту ложатся туда, где действительно стоят - на 2 и 3.
# ---------------------------------------------------------------------------
$r12 = New-TestRun
New-TestEvents (Join-Path $r12.Dir 'events.tsv') @(
    @{ Name = 'ALFunctionStart'; Ms = 1;  Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '';  Stmt = '';                  Rec = 1 },
    @{ Name = 'ALStatement';     Ms = 1;  Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '2'; Stmt = 'Tick;';             Rec = 2 },
    @{ Name = 'ALStatement';     Ms = 4;  Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '2'; Stmt = 'Done := Done + 1;'; Rec = 3 },
    @{ Name = 'ALFunctionStop';  Ms = 10; Ot = 5; Oi = 110200; Fn = 'MyFunc'; Line = '';  Stmt = '';                  Rec = 4 }
)
$code12 = Invoke-Rebuild -RunDir $r12.Dir -Src $r12.Src
Test-Value 'калибровка не сошлась: код возврата' 0 $code12
Test-Value 'калибровка не сошлась: строки по тексту' '2,3' ((Get-MetricLines $r12.Dir) -join ',')
Remove-Item -LiteralPath $r12.Dir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# вердикт
# ---------------------------------------------------------------------------
Write-Host ('пройдено {0} из {1}' -f $script:Passed, $script:Total)
foreach ($f in $script:Fails) { Write-Host ('  ' + $f) -ForegroundColor Red }
if ($script:Fails.Count -gt 0) { exit 1 }
Write-Host 'ядро разбора трассировки: без расхождений' -ForegroundColor Green
exit 0
