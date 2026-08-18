#Requires -Version 5.1
<#
.SYNOPSIS
    Прогон правил линтера на синтетическом объекте с заранее известным ответом.

.DESCRIPTION
    Массовый прогон линтера требует выгрузки объектов установки, и на чистой копии
    репозитория правила не проверялись ничем. Между тем врут они дорого: совет
    «вынести до цикла», выданный с уверенностью «высокая», разработчик выполняет —
    и получает одну учётную группу на все записи.

    Поэтому объект собирается прямо здесь: дамп исходников (.alsrc) и текстовый
    экспорт для таблицы символов. Таблицы в примерах штатные (18 Customer,
    92 Customer Posting Group, 348 Dimension, 37 Sales Line), номера объектов — из
    диапазона 50000, к какой-либо установке отношения не имеющие.

    Судит сама: первая строка вердикта — «пройдено N из M», код возврата ненулевой
    при расхождении. Ни базы, ни трассировки, ни выгрузки не требует.

.EXAMPLE
    pwsh scripts/Test-PerfLint.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Total  = 0
$script:Fails  = New-Object System.Collections.ArrayList

function Test-Value {
    param([string]$What, $Expect, $Got)
    $script:Total++
    if ("$Expect" -eq "$Got") { $script:Passed++; return }
    [void]$script:Fails.Add(('{0}: ждали <{1}>, получили <{2}>' -f $What, $Expect, $Got))
}

$U   = New-Object System.Text.UTF8Encoding($false)
$TAB = [char]9
$root = Join-Path ([System.IO.Path]::GetTempPath()) ('lp-lint-' + [guid]::NewGuid().ToString('N'))
$src  = Join-Path $root 'src'
$base = Join-Path $root 'base'
New-Item -ItemType Directory -Path $src -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $base 'Codeunits') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $base 'Reports')   -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $base 'Pages')     -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $base 'XMLports')  -Force | Out-Null

# ---------------------------------------------------------------------------
# Дамп исходников: столбец 0 занят — заголовок функции, отступ — тело.
# ---------------------------------------------------------------------------
# Поле берётся БЕЗ квалификатора, через WITH, — ровно так это и пишут в рабочем
# коде, и именно такой корень прежний разбор считал константой.
$cu = @(
    'ПолеПеребираемойЗаписи()'                            # 1
    '  WITH Buffer DO'                                    # 2
    '    IF FIND(''-'') THEN'                             # 3
    '      REPEAT'                                        # 4
    '        CustPostingGr.GET("Gen. Bus. Posting Group");' # 5
    '      UNTIL NEXT = 0;'                               # 6
    'СчётчикЦикла()'                                      # 7
    '  FOR k := 1 TO 10 DO BEGIN'                         # 8
    '    Dimension.GET(DimArr[k]);'                       # 9
    '  END;'                                              # 10
    'ДействительноИнвариантный()'                         # 11
    '  IF Buffer.FIND(''-'') THEN'                        # 12
    '    REPEAT'                                          # 13
    '      Cust.GET(ConstCode);'                          # 14
    '    UNTIL Buffer.NEXT = 0;'                          # 15
    'ГорячаяБезЦикла(No : Code[20])'                      # 16
    '  Cust.GET(No);'                                     # 17
)
[System.IO.File]::WriteAllText((Join-Path $src '5_50000.al'), (($cu -join "`r`n") + "`r`n"), $U)

$rep = @(
    'Sales Invoice Header - OnAfterGetRecord()'                  # 1
    '  IF NOT Cust.GET("Bill-to Customer No.") THEN'             # 2
    '    Cust.INIT;'                                             # 3
)
[System.IO.File]::WriteAllText((Join-Path $src '3_50001.al'), (($rep -join "`r`n") + "`r`n"), $U)

# У страницы безымянный вызов адресует Rec: без этого в тексте находки выходила
# дыра — «CALCFIELDS(...) по  в цикле» с двойным пробелом.
$pg = @(
    'OnAfterGetRecord()'                        # 1
    '  CALCFIELDS("Balance (LCY)");'            # 2
)
[System.IO.File]::WriteAllText((Join-Path $src '8_50002.al'), (($pg -join "`r`n") + "`r`n"), $U)

# В XMLport'е к имени триггера добавлено направление: «X - Export::OnAfterGetRecord».
# Без его учёта разбор молча не срабатывал, и включение типа 6 было пустой строкой.
$xp = @(
    'Sales Line - Export::OnAfterGetRecord()'   # 1
    '  IF NOT Cust.GET("Sell-to Customer No.") THEN' # 2
    '    Cust.INIT;'                            # 3
)
[System.IO.File]::WriteAllText((Join-Path $src '6_50003.al'), (($xp -join "`r`n") + "`r`n"), $U)

$idx = @(
    (@('ObjectType','TypeName','ObjectId','Name','Lines','Bytes','Compiled','Date','Time','VersionList','Hash') -join $TAB)
    (@('5','Codeunit','50000','Проверка правил','17','420','1','19.08.26','12:00:00','TEST','0') -join $TAB)
    (@('3','Report','50001','Проверка правил отчёта','3','120','1','19.08.26','12:00:00','TEST','0') -join $TAB)
    (@('8','Page','50002','Проверка правил страницы','2','80','1','19.08.26','12:00:00','TEST','0') -join $TAB)
    (@('6','XMLport','50003','Проверка правил XMLport','3','120','1','19.08.26','12:00:00','TEST','0') -join $TAB)
)
[System.IO.File]::WriteAllText((Join-Path $src 'index.tsv'), (($idx -join "`r`n") + "`r`n"), $U)

# ---------------------------------------------------------------------------
# Экспорт для таблицы символов: типы переменных берутся отсюда.
# ---------------------------------------------------------------------------
$cuExp = @(
    'OBJECT Codeunit 50000 Проверка правил'
    '{'
    '  OBJECT-PROPERTIES'
    '  {'
    '    Date=19.08.26;'
    '    Time=12:00:00;'
    '    Modified=Yes;'
    '    Version List=TEST;'
    '  }'
    '  PROPERTIES'
    '  {'
    '  }'
    '  CODE'
    '  {'
    '    VAR'
    '      Cust@1000000000 : Record 18;'
    '      CustPostingGr@1000000001 : Record 92;'
    '      Dimension@1000000002 : Record 348;'
    '      Buffer@1000000003 : Record 37;'
    '      DimArr@1000000004 : ARRAY [10] OF Code[20];'
    '      ConstCode@1000000005 : Code[20];'
    '      k@1000000006 : Integer;'
    ''
    '    BEGIN'
    '    END.'
    '  }'
    '}'
)
[System.IO.File]::WriteAllText((Join-Path $base 'Codeunits\c50000 - Проверка правил.txt'),
    (($cuExp -join "`r`n") + "`r`n"), $U)

$repExp = @(
    'OBJECT Report 50001 Проверка правил отчёта'
    '{'
    '  OBJECT-PROPERTIES'
    '  {'
    '    Date=19.08.26;'
    '    Time=12:00:00;'
    '    Modified=Yes;'
    '    Version List=TEST;'
    '  }'
    '  PROPERTIES'
    '  {'
    '  }'
    '  CODE'
    '  {'
    '    VAR'
    '      Cust@1000000000 : Record 18;'
    ''
    '    BEGIN'
    '    END.'
    '  }'
    '}'
)
[System.IO.File]::WriteAllText((Join-Path $base 'Reports\r50001 - Проверка правил отчёта.txt'),
    (($repExp -join "`r`n") + "`r`n"), $U)

$pgExp = @(
    'OBJECT Page 50002 Проверка правил страницы'
    '{'
    '  OBJECT-PROPERTIES'
    '  {'
    '    Date=19.08.26;'
    '    Time=12:00:00;'
    '    Modified=Yes;'
    '    Version List=TEST;'
    '  }'
    '  PROPERTIES'
    '  {'
    '  }'
    '  CODE'
    '  {'
    '    VAR'
    '      Dummy@1000000000 : Integer;'
    ''
    '    BEGIN'
    '    END.'
    '  }'
    '}'
)
[System.IO.File]::WriteAllText((Join-Path $base 'Pages\p50002 - Проверка правил страницы.txt'),
    (($pgExp -join "`r`n") + "`r`n"), $U)

$xpExp = @(
    'OBJECT XMLport 50003 Проверка правил XMLport'
    '{'
    '  OBJECT-PROPERTIES'
    '  {'
    '    Date=19.08.26;'
    '    Time=12:00:00;'
    '    Modified=Yes;'
    '    Version List=TEST;'
    '  }'
    '  PROPERTIES'
    '  {'
    '  }'
    '  CODE'
    '  {'
    '    VAR'
    '      Cust@1000000000 : Record 18;'
    ''
    '    BEGIN'
    '    END.'
    '  }'
    '}'
)
[System.IO.File]::WriteAllText((Join-Path $base 'XMLports\x50003 - Проверка правил XMLport.txt'),
    (($xpExp -join "`r`n") + "`r`n"), $U)

# Метрики нужны одному случаю: цикл, которого в тексте нет, а по числу попаданий он есть.
$met = @(
    (@('ObjectType','ObjectId','LineNo','Hits','TotalMs','SelfMs','MinMs','MaxMs','SqlCount','SqlMs','Stmts') -join $TAB)
    (@('5','50000','17','500','600','600','1','10','500','400','1') -join $TAB)
)
$metFile = Join-Path $root 'lines.tsv'
[System.IO.File]::WriteAllText($metFile, (($met -join "`r`n") + "`r`n"), $U)

# ---------------------------------------------------------------------------
# Прогон
# ---------------------------------------------------------------------------
function Invoke-Lint {
    param([int]$Ot, [int]$Oi, [string]$Metrics)
    $out = Join-Path $root ('out-{0}-{1}.tsv' -f $Ot, $Oi)
    # Порог задаём явно: без метрик он по умолчанию Info, с метриками Medium, и
    # прогон с метриками молча терял бы находки, которые проверяются здесь.
    $ba = @{
        ObjectType = $Ot; ObjectId = $Oi
        SourceRoot = $src; BaseRoot = $base
        OutFile    = $out; Quiet = $true
        MinSeverity = 'Info'
    }
    if ($Metrics) { $ba['MetricsFile'] = $Metrics }
    try { & (Join-Path $PSScriptRoot 'Invoke-PerfLint.ps1') @ba *>$null } catch { }
    if (-not (Test-Path -LiteralPath $out)) { return @() }
    $rows = @()
    $lines = [System.IO.File]::ReadAllLines($out, [System.Text.Encoding]::UTF8)
    if ($lines.Length -lt 2) { return @() }
    $h = $lines[0] -split "`t"
    for ($i = 1; $i -lt $lines.Length; $i++) {
        if (-not $lines[$i]) { continue }
        $c = $lines[$i] -split "`t"
        $o = @{}
        for ($j = 0; $j -lt $h.Length -and $j -lt $c.Length; $j++) { $o[$h[$j]] = $c[$j] }
        $rows += [pscustomobject]$o
    }
    return $rows
}

function Get-Msg {
    <#  Сообщение правила на строке; '' — если находки нет.  #>
    param($Rows, [int]$Line, [string]$Rule)
    foreach ($r in $Rows) {
        if ([int]$r.'Строка' -eq $Line -and $r.RuleID -eq $Rule) { return $r.'Сообщение' }
    }
    return ''
}

$rowsCu  = Invoke-Lint -Ot 5 -Oi 50000 -Metrics $metFile
$rowsRep = Invoke-Lint -Ot 3 -Oi 50001
$rowsPg  = Invoke-Lint -Ot 8 -Oi 50002
$rowsXp  = Invoke-Lint -Ot 6 -Oi 50003

# ---------------------------------------------------------------------------
# 1. Аргумент — голое поле перебираемой записи. Меняется с каждой записью,
#    выносить нельзя: получится одна учётная группа на все записи.
# ---------------------------------------------------------------------------
$m1 = Get-Msg $rowsCu 5 '111'
Test-Value 'поле перебираемой записи: находка есть'   $true  ($m1.Length -gt 0)
Test-Value 'поле перебираемой записи: не инвариантно' $false ($m1 -match 'инвариантно')
Test-Value 'поле перебираемой записи: уверенность'    $true  ($m1 -match 'уверенность: низкая')

# ---------------------------------------------------------------------------
# 2. Счётчик FOR меняется каждый виток. В присваивания он не попадает: их
#    regex заякорен на начало строки, а строка начинается со слова FOR.
# ---------------------------------------------------------------------------
$m2 = Get-Msg $rowsCu 9 '111'
Test-Value 'счётчик FOR: находка есть'   $true  ($m2.Length -gt 0)
Test-Value 'счётчик FOR: не инвариантно' $false ($m2 -match 'инвариантно')

# ---------------------------------------------------------------------------
# 3. Обратная сторона: настоящий инвариант обязан остаться инвариантом, иначе
#    осторожность съела бы весь смысл правила.
# ---------------------------------------------------------------------------
$m3 = Get-Msg $rowsCu 14 '111'
Test-Value 'настоящий инвариант: находка есть'  $true ($m3.Length -gt 0)
Test-Value 'настоящий инвариант: так и сказано' $true ($m3 -match 'инвариантно')
Test-Value 'настоящий инвариант: уверенность'   $true ($m3 -match 'уверенность: высокая')

# ---------------------------------------------------------------------------
# 4. Цикла в тексте нет вовсе — он виден только по числу попаданий в замере.
#    Границ у такого цикла нет, доказывать инвариантность нечем.
# ---------------------------------------------------------------------------
$m4 = Get-Msg $rowsCu 17 '111'
Test-Value 'цикл только по замеру: находка есть'   $true  ($m4.Length -gt 0)
Test-Value 'цикл только по замеру: не инвариантно' $false ($m4 -match 'инвариантно')

# ---------------------------------------------------------------------------
# 5. Триггер, который платформа зовёт на каждой записи: синтаксического цикла
#    нет, аргумент — поле датаитема, выносить некуда и нельзя.
# ---------------------------------------------------------------------------
$m5 = Get-Msg $rowsRep 2 '111'
Test-Value 'триггер на каждой записи: находка есть'   $true  ($m5.Length -gt 0)
Test-Value 'триггер на каждой записи: не инвариантно' $false ($m5 -match 'инвариантно')

# ---------------------------------------------------------------------------
# 6. Страница: безымянный вызов адресует Rec. Без этого в тексте находки
#    оставалась дыра на месте записи — «CALCFIELDS(...) по  в цикле».
# ---------------------------------------------------------------------------
$m6 = Get-Msg $rowsPg 2 '110'
Test-Value 'страница: находка есть'      $true  ($m6.Length -gt 0)
Test-Value 'страница: запись названа'    $true  ($m6 -match 'по Rec в цикле')
Test-Value 'страница: дыры в тексте нет' $false ($m6 -match 'по\s\sв цикле')

# ---------------------------------------------------------------------------
# 7. XMLport: в имени триггера стоит направление, «X - Export::OnAfterGetRecord».
#    Без его учёта тип 6 был включён впустую — ни одной «цикловой» находки.
# ---------------------------------------------------------------------------
$m7 = Get-Msg $rowsXp 2 '111'
Test-Value 'XMLport: цикл распознан'     $true  ($m7.Length -gt 0)
Test-Value 'XMLport: не инвариантно'     $false ($m7 -match 'инвариантно')

# ---------------------------------------------------------------------------
# вердикт
# ---------------------------------------------------------------------------
Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ('пройдено {0} из {1}' -f $script:Passed, $script:Total)
foreach ($f in $script:Fails) { Write-Host ('  ' + $f) -ForegroundColor Red }
if ($script:Fails.Count -gt 0) { exit 1 }
Write-Host 'правила линтера: без расхождений' -ForegroundColor Green
exit 0
