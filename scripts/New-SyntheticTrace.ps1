#Requires -Version 5.1
<#
.SYNOPSIS
    Генератор синтетической трассировки C/AL: events.tsv в формате Collect-AlTrace.ps1
    плюс expected.tsv — эталонные метрики строк, посчитанные независимо от разбора.

.DESCRIPTION
    Сбор трассировки на живой базе требует прав администратора и перезапуска инстанции.
    Чтобы арифметика разбора (Lib-AlTrace.ps1) проверялась до того, как трассировку
    вообще удастся снять, здесь эмулируется поток событий по РЕАЛЬНОМУ листингу объекта
    из дампа .alsrc: номера строк, тексты операторов, вложенные вызовы и обращения к SQL
    берутся из настоящего кода, а длительности назначаются детерминированно.

    Случайности нет вовсе. Собственное время k-го оператора = -BaseTicks +
    (k mod -CycleLen) * -StepTicks: арифметическая последовательность, зациклённая,
    чтобы соседние операторы стоили по-разному, а суммарное время не убегало. Поэтому
    ожидаемые Hits/Total/Self считаются заранее, ПО ЗАМЫСЛУ, и ложатся в expected.tsv;
    разбор же восстанавливает их только из меток времени и порядка событий. Сходятся
    два независимых пути — значит арифметика верна.

    Что эмулируется:
      - вход в функцию (роль ALFunctionStart) и выход (ALFunctionStop) с постоянными
        накладными расходами рамки -FrameOverheadTicks с каждой стороны; эпилог
        рамки эталон относит на ПОСЛЕДНЮЮ строку функции — между ней и событием
        выхода событий нет, и разбор не может отнести это время никуда больше;
      - оператор (ALStatement) на каждый оператор функции; текст СОВПАДАЕТ с текстом
        строки листинга, многострочный оператор склеивается в один текст, как это
        делает платформа (по балансу скобок);
      - вложенные вызовы: имя функции этого же объекта -> рамка на нём же; вызов вида
        Переменная.Метод(...) -> рамка на «чужом» объекте (-ForeignObjectType/Id);
        встроенные методы записи (GET, FINDSET, MODIFY...) рамкой не считаются;
      - цикл REPEAT..UNTIL: тело повторяется -Iterations раз, строка UNTIL получает по
        событию на итерацию, строка REPEAT — одно;
      - SQL: на операторах с GET/FINDSET/CALCSUMS/MODIFY и т.п. пара событий
        <операция>:Start / <операция>:Stop с текстом запроса в сыром столбце;
      - фоновая сессия: события другой сессии, вклинивающиеся в поток главной, чтобы
        проверить, что стеки вызовов разбираются по сессиям раздельно. Фоновые события
        эмитятся в ТУ ЖЕ метку времени (нулевая длительность) — иначе они сдвинули бы
        часы и сломали эталонную арифметику главной сессии.

    Номер строки в событии = номер строки листинга + -LineNumberBias. По умолчанию
    -1, то есть правильное смещение при разборе +1 — ровно то, что должна найти
    самокалибровка Resolve-AlLineOffset. Другое значение параметра позволяет
    проверить, что калибровка находит и 0, и -1.

    Ключ -Verify прогоняет разбор по собственному выводу и печатает таблицу сверки
    с expected.tsv плюс отдельные проверки: цикл, вложенный вызов, калибровка,
    резервный маппинг, изоляция сессий, сводка прогона.

    Кодировки: events.tsv пишется UTF-8 БЕЗ BOM с CRLF — байт в байт как его пишет
    Collect-AlTrace.ps1, чтобы разбор проверялся на реальном формате; expected.tsv —
    UTF-8 с BOM.

.PARAMETER ObjectType
    Тип объекта (5 — Codeunit).

.PARAMETER ObjectId
    Номер объекта.

.PARAMETER FunctionName
    Функция, выполнение которой эмулируется. По умолчанию — самая длинная в объекте.

.PARAMETER OutDir
    Куда писать. По умолчанию <задача>\out\synth.

.PARAMETER Iterations
    Число итераций цикла.

.PARAMETER LoopBodyMax
    Сколько первых операторов тела цикла выполнять (ограничивает размер файла).

.PARAMETER CalleeLines
    Сколько первых операторов выполняет вызванная функция.

.PARAMETER RecursionDepth
    Глубина рекурсии: первый оператор тела цикла вызывает свою же функцию N раз
    вложенно. 0 — рекурсии нет.

.PARAMETER LineNumberBias
    Сдвиг номера строки в событии относительно листинга.

.PARAMETER CompositeProbe
    Составная строка: второй оператор тела цикла эмитится с номером строки ПЕРВОГО.
    Так выглядит в потоке событий составной оператор — два разных текста на одной
    строке; проверяет, что метрики агрегируются суммой.

.PARAMETER Verify
    Прогнать разбор по своему выводу и напечатать сверку с эталоном.

.EXAMPLE
    .\New-SyntheticTrace.ps1 -Verify
    Синтетика по самой длинной функции Codeunit 80 и сверка с эталоном.

.EXAMPLE
    .\New-SyntheticTrace.ps1 -FunctionName PostItemJnlLine -Iterations 50 -Verify
#>
[CmdletBinding()]
param(
    [int]    $ObjectType = 5,
    [int]    $ObjectId   = 80,
    [string] $FunctionName,
    [string] $OutDir,
    [string] $SourceRoot,
    [int]    $Iterations         = 1000,
    [int]    $LoopBodyMax        = 8,
    [int]    $CalleeLines        = 2,
    [int]    $MaxCallDepth       = 1,
    [int]    $RecursionDepth     = 0,
    [int]    $LineNumberBias     = -1,
    [switch] $CompositeProbe,
    [string] $SessionId          = '42',
    [string] $NoiseSessionId     = '77',
    [int]    $NoiseEvery         = 25,
    [switch] $NoNoise,
    [int]    $BaseTicks          = 1000,
    [int]    $StepTicks          = 100,
    [int]    $CycleLen           = 16,
    [int]    $FrameOverheadTicks = 50,
    [int]    $SqlBaseTicks       = 3000,
    [int]    $SqlStepTicks       = 250,
    [int]    $ForeignObjectType  = 1,
    [int]    $ForeignObjectId    = 1235,
    [switch] $Verify,
    [switch] $PassThru,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Lib-AlListing.ps1')

# ---------------------------------------------------------------------------
# разбор текста C/AL
# ---------------------------------------------------------------------------

$script:BuiltinMethods = @{}
foreach ($m in @(
    'GET','FIND','FINDSET','FINDFIRST','FINDLAST','NEXT','INSERT','MODIFY','DELETE',
    'DELETEALL','MODIFYALL','INIT','RESET','SETRANGE','SETFILTER','GETFILTER','GETFILTERS',
    'SETCURRENTKEY','CALCFIELDS','CALCSUMS','COUNT','COUNTAPPROX','ISEMPTY','VALIDATE',
    'TESTFIELD','TRANSFERFIELDS','COPY','COPYFILTERS','FIELDNO','FIELDCAPTION','TABLECAPTION',
    'RECORDID','SETPOSITION','GETPOSITION','LOCKTABLE','MARK','MARKEDONLY','CLEARMARKS',
    'ASCENDING','RENAME','SETAUTOCALCFIELDS','CURRENTKEY','SETVIEW','GETVIEW','CHANGECOMPANY'
)) { $script:BuiltinMethods[$m] = $true }

$script:SqlMethods = @{}
foreach ($m in @('GET','FIND','FINDSET','FINDFIRST','FINDLAST','NEXT','CALCSUMS','CALCFIELDS',
                 'COUNT','ISEMPTY','INSERT','MODIFY','DELETE','DELETEALL','MODIFYALL')) {
    $script:SqlMethods[$m] = $true
}
$script:SqlWriteMethods = @{}
foreach ($m in @('INSERT','MODIFY','DELETE','DELETEALL','MODIFYALL')) { $script:SqlWriteMethods[$m] = $true }

function Remove-AlComment {
    <#
    .SYNOPSIS
        Отрезает хвостовой комментарий, не трогая '//' внутри строковой константы.
    #>
    param([string] $Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $inStr = $false
    $n = $Text.Length
    for ($i = 0; $i -lt $n; $i++) {
        $c = $Text[$i]
        if ($inStr) { if ($c -eq "'") { $inStr = $false }; continue }
        if ($c -eq "'") { $inStr = $true; continue }
        if ($c -eq '/' -and ($i + 1) -lt $n -and $Text[$i + 1] -eq '/') { return $Text.Substring(0, $i).Trim() }
    }
    return $Text.Trim()
}

function Get-AlParenBalance {
    <#
    .SYNOPSIS
        Баланс круглых скобок вне строковых констант: >0 — оператор продолжается на следующей строке.
    #>
    param([string] $Text)
    $b = 0
    $inStr = $false
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $c = $Text[$i]
        if ($inStr) { if ($c -eq "'") { $inStr = $false }; continue }
        if ($c -eq "'") { $inStr = $true; continue }
        if     ($c -eq '(') { $b++ }
        elseif ($c -eq ')') { $b-- }
    }
    return $b
}

function New-SynthStatementList {
    <#
    .SYNOPSIS
        Разбивает строки листинга на операторы: склеивает продолжения по балансу скобок,
        находит вызовы и обращения к базе.

    .DESCRIPTION
        Возвращает на оператор: первую строку (номер как в листинге), склеенный текст,
        признак обращения к SQL, имя и вид вызываемой функции. Вид вызова:
        Internal — функция того же объекта, External — Переменная.Метод(), встроенные
        методы записи вызовом не считаются (у них своих событий входа/выхода нет).
    #>
    param(
        [Parameter(Mandatory)][object[]] $Listing,
        [Parameter(Mandatory)][int]      $From,
        [Parameter(Mandatory)][int]      $To,
        [hashtable] $FuncNames,
        [int]       $MaxJoin = 6
    )

    $code = New-Object System.Collections.Generic.List[object]
    foreach ($l in $Listing) {
        if ($l.LineNo -ge $From -and $l.LineNo -le $To -and $l.Kind -eq 'Code') { [void]$code.Add($l) }
    }

    $rxCall    = New-Object System.Text.RegularExpressions.Regex('([A-Za-z_][A-Za-z0-9_]*)\s*\(', 'Compiled')
    $rxMethod  = New-Object System.Text.RegularExpressions.Regex('([A-Za-z_][A-Za-z0-9_]*)\s*\.\s*([A-Za-z_][A-Za-z0-9_]*)', 'Compiled')

    $res = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $code.Count) {
        $first = $code[$i]
        $t   = Remove-AlComment $first.Text
        $bal = Get-AlParenBalance $t
        $parts = New-Object System.Collections.Generic.List[string]
        [void]$parts.Add($t)
        $j = $i + 1
        while ($bal -gt 0 -and $j -lt $code.Count -and ($j - $i) -lt $MaxJoin) {
            $t2 = Remove-AlComment $code[$j].Text
            if ($t2.Length -gt 0) {
                [void]$parts.Add($t2)
                $bal += Get-AlParenBalance $t2
            }
            $j++
        }
        $text = ($parts -join ' ').Trim()
        $i = $j
        if ($text.Length -eq 0) { continue }

        # обращение к базе и вызов
        $isSql = $false; $sqlOp = ''
        $callee = ''; $calleeKind = ''
        foreach ($m in $rxMethod.Matches($text)) {
            $meth = $m.Groups[2].Value.ToUpperInvariant()
            if ($script:SqlMethods.ContainsKey($meth) -and -not $isSql) {
                $isSql = $true
                if ($script:SqlWriteMethods.ContainsKey($meth)) { $sqlOp = 'ExecuteNonQuery' } else { $sqlOp = 'ExecuteReader' }
            }
            if ($callee.Length -eq 0 -and -not $script:BuiltinMethods.ContainsKey($meth)) {
                # вызов метода переменной: рамка на чужом объекте
                $after = $m.Index + $m.Length
                if ($after -lt $text.Length -and $text.Substring($after).TrimStart().StartsWith('(')) {
                    $callee = $m.Groups[1].Value + '.' + $m.Groups[2].Value
                    $calleeKind = 'External'
                }
            }
        }
        if ($FuncNames) {
            foreach ($m in $rxCall.Matches($text)) {
                $name = $m.Groups[1].Value
                if (-not $FuncNames.ContainsKey($name)) { continue }
                if ($m.Index -gt 0 -and $text[$m.Index - 1] -eq '.') { continue }
                $callee = $name; $calleeKind = 'Internal'
                break
            }
        }

        [void]$res.Add([pscustomobject]@{
            Line       = $first.LineNo
            EmitLine   = $first.LineNo
            Text       = $text
            LineCount  = ($j - $i)
            IsSql      = $isSql
            SqlOp      = $sqlOp
            Callee     = $callee
            CalleeKind = $calleeKind
        })
    }
    return $res
}

# ---------------------------------------------------------------------------
# эмиссия событий
# ---------------------------------------------------------------------------

function Step-SynthClock {
    param([int64] $Ticks)
    $script:Now += $Ticks
}

function Add-SynthRow {
    <#
    .SYNOPSIS
        Пишет строку events.tsv в порядке колонок Collect-AlTrace.ps1.
    #>
    param(
        [int]    $EventId,
        [string] $Name,
        [string] $Session,
        [string] $ObjType = '',
        [string] $ObjId   = '',
        [string] $Func    = '',
        [string] $Line    = '',
        [string] $Stmt    = '',
        [string] $Raw     = '',
        [string] $Thread  = '4120',
        [switch] $NoNoiseHook
    )
    $script:Rec++
    $tab = [char]9
    [void]$script:Rows.Add(
        [string]$EventId + $tab + $Name + $tab + $script:Now.ToString() + $tab + $Session + $tab +
        $ObjType + $tab + $ObjId + $tab + $Func + $tab + $Line + $tab + $Stmt + $tab +
        '4' + $tab + $script:Rec.ToString() + $tab + $Thread + $tab + $Raw)

    if (-not $NoNoiseHook -and $script:NoiseOn -and $Session -eq $script:MainSession) {
        $script:MainRows++
        if (($script:MainRows % $script:NoiseEvery) -eq 0) { Invoke-SynthNoise }
    }
}

function Invoke-SynthNoise {
    <#
    .SYNOPSIS
        Одно событие фоновой сессии — в текущую метку времени, без сдвига часов.
    #>
    $s = $script:NoiseSteps[$script:NoiseIx % $script:NoiseSteps.Count]
    $script:NoiseIx++
    Add-SynthRow -EventId $s.Id -Name $s.Name -Session $script:NoiseSession `
                 -ObjType ([string]$script:ObjTypeS) -ObjId ([string]$script:ObjIdS) `
                 -Func $s.Func -Line $s.Line -Stmt $s.Text -Thread '4121' -NoNoiseHook
}

function Update-SynthExpected {
    <#
    .SYNOPSIS
        Копит ЭТАЛОННЫЕ метрики строки по замыслу, не по меткам времени.
    #>
    param(
        [int] $Line, [string] $Func, [int64] $Incl, [int64] $Self,
        [int64] $Child, [int] $SqlCount, [int64] $SqlTicks
    )
    $b = $script:Expected[$Line]
    if ($null -eq $b) {
        $b = @{
            Line = $Line; Func = $Func; Hits = 0L; Total = 0L; Self = 0L; Child = 0L
            SqlCount = 0L; SqlTicks = 0L; Min = [int64]::MaxValue; Max = 0L
        }
        $script:Expected[$Line] = $b
    }
    $b.Hits++
    $b.Total    += $Incl
    $b.Self     += $Self
    $b.Child    += $Child
    $b.SqlCount += $SqlCount
    $b.SqlTicks += $SqlTicks
    if ($Incl -lt $b.Min) { $b.Min = $Incl }
    if ($Incl -gt $b.Max) { $b.Max = $Incl }
}

function Invoke-SynthSql {
    <#
    .SYNOPSIS
        Пара событий обращения к базе; возвращает её длительность в тиках.
    #>
    param([string] $Session, [string] $Op, [string] $Thread)

    $script:SqlIx++
    $dur = [int64]($script:SqlBase + ($script:SqlIx % 8) * $script:SqlStep)
    $id  = 5
    if ($Op -eq 'ExecuteNonQuery') { $id = 3 }
    # точка с запятой внутри запроса — намеренно: сырой столбец разбирается «до конца строки»
    $raw = 'userName=DD; sqlStatement=SET NOCOUNT ON; SELECT TOP 1 * FROM [dbo].[NAV$Sales Line] WHERE ([Entry No_]=@0)'

    Add-SynthRow -EventId $id       -Name ('Sql:' + $Op + ':Start') -Session $Session -Raw $raw -Thread $Thread
    Step-SynthClock $dur
    Add-SynthRow -EventId ($id + 1) -Name ('Sql:' + $Op + ':Stop')  -Session $Session -Raw $raw -Thread $Thread
    return $dur
}

function Invoke-SynthStatement {
    <#
    .SYNOPSIS
        Один оператор: событие, собственное время, при необходимости SQL и вложенный вызов.
    #>
    param(
        [string] $Session, [int] $ObjType, [int] $ObjId, [string] $Func,
        $Stmt, [int] $Depth, [int] $RecLeft, [string] $Thread, [int64] $ExtraSelf = 0
    )

    $t0 = $script:Now
    Add-SynthRow -EventId 403 -Name 'ALStatement' -Session $Session `
                 -ObjType ([string]$ObjType) -ObjId ([string]$ObjId) -Func $Func `
                 -Line ([string]($Stmt.EmitLine + $script:Bias)) -Stmt $Stmt.Text -Thread $Thread

    $script:StmtIx++
    $self = [int64]($script:BaseTicks + ($script:StmtIx % $script:CycleLen) * $script:StepTicks)
    $half = [int64][math]::Floor($self / 2)
    Step-SynthClock $half

    $sqlTicks = 0L
    $sqlCount = 0
    if ($Stmt.IsSql) {
        $sqlTicks = Invoke-SynthSql -Session $Session -Op $Stmt.SqlOp -Thread $Thread
        $sqlCount = 1
    }

    $child = 0L
    if ($RecLeft -gt 0 -and $script:RecTrigger -gt 0 -and $Stmt.EmitLine -eq $script:RecTrigger) {
        $child += Invoke-SynthFrame -Session $Session -ObjType $script:ObjTypeS -ObjId $script:ObjIdS `
                                    -Func $script:TargetFunc -Stmts $script:RecStmts `
                                    -Depth ($Depth + 1) -RecLeft ($RecLeft - 1) -Thread $Thread
    }
    elseif ($Depth -lt $script:MaxDepth -and $Stmt.Callee.Length -gt 0) {
        if ($Stmt.CalleeKind -eq 'Internal') {
            $steps = @(Get-SynthCalleeSteps -Name $Stmt.Callee)
            if ($steps.Count -gt 0) {
                $child += Invoke-SynthFrame -Session $Session -ObjType $script:ObjTypeS -ObjId $script:ObjIdS `
                                            -Func $Stmt.Callee -Stmts $steps -Depth ($Depth + 1) `
                                            -RecLeft 0 -Thread $Thread
            }
        }
        else {
            $child += Invoke-SynthFrame -Session $Session -ObjType $script:ForeignType -ObjId $script:ForeignId `
                                        -Func $Stmt.Callee -Stmts $script:ForeignStmts -Depth ($Depth + 1) `
                                        -RecLeft 0 -Thread $Thread
        }
    }

    # ExtraSelf — эпилог рамки: между последним оператором функции и событием выхода
    # своего события нет, и это время неизбежно ложится на последнюю строку. Эталон
    # обязан моделировать это так же, иначе он расходится с разбором на ровном месте.
    Step-SynthClock ($self - $half + $ExtraSelf)

    if ($Session -eq $script:MainSession -and $ObjType -eq $script:ObjTypeS -and $ObjId -eq $script:ObjIdS) {
        Update-SynthExpected -Line $Stmt.EmitLine -Func $Func -Incl ($script:Now - $t0) `
                             -Self ($self + $sqlTicks + $ExtraSelf) -Child $child `
                             -SqlCount $sqlCount -SqlTicks $sqlTicks
    }
}

function Invoke-SynthFrame {
    <#
    .SYNOPSIS
        Рамка вызова: вход, операторы, выход. Возвращает длительность рамки в тиках.
    #>
    param(
        [string] $Session, [int] $ObjType, [int] $ObjId, [string] $Func,
        [object[]] $Stmts, [int] $Depth, [int] $RecLeft, [string] $Thread
    )

    $t0 = $script:Now
    Add-SynthRow -EventId 400 -Name 'ALFunctionStart' -Session $Session `
                 -ObjType ([string]$ObjType) -ObjId ([string]$ObjId) -Func $Func -Thread $Thread
    Step-SynthClock $script:FrameOverhead

    for ($i = 0; $i -lt $Stmts.Count; $i++) {
        $extra = 0L
        if ($i -eq ($Stmts.Count - 1)) { $extra = [int64]$script:FrameOverhead }
        Invoke-SynthStatement -Session $Session -ObjType $ObjType -ObjId $ObjId -Func $Func `
                              -Stmt $Stmts[$i] -Depth $Depth -RecLeft $RecLeft -Thread $Thread `
                              -ExtraSelf $extra
    }

    Add-SynthRow -EventId 401 -Name 'ALFunctionStop' -Session $Session `
                 -ObjType ([string]$ObjType) -ObjId ([string]$ObjId) -Func $Func -Thread $Thread
    return ($script:Now - $t0)
}

function Get-SynthCalleeSteps {
    <#
    .SYNOPSIS
        Первые операторы вызванной функции того же объекта; результат кэшируется.
    #>
    param([string] $Name)

    if ($script:CalleeCache.ContainsKey($Name)) { return $script:CalleeCache[$Name] }
    $take = @()
    $fn = $script:FuncByName[$Name]
    if ($fn) {
        $lst = New-SynthStatementList -Listing $script:Listing -From $fn.FirstLine -To $fn.LastLine -FuncNames $null
        $n = [math]::Min($script:CalleeLines, $lst.Count)
        for ($i = 0; $i -lt $n; $i++) { $take += $lst[$i] }
    }
    $script:CalleeCache[$Name] = $take
    return $take
}

# ---------------------------------------------------------------------------
# подготовка
# ---------------------------------------------------------------------------

if (-not $OutDir) { $OutDir = Join-Path (Split-Path -Parent $PSScriptRoot) 'out\synth' }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$info = Get-AlObjectInfo -ObjectType $ObjectType -ObjectId $ObjectId -SourceRoot $SourceRoot
$listing = Get-AlListing -ObjectType $ObjectType -ObjectId $ObjectId -SourceRoot $SourceRoot
$fmap    = Get-AlFunctionMap -Listing $listing

$funcSet    = @{}
$funcByName = @{}
foreach ($f in $fmap) { $funcSet[$f.Name] = $true; $funcByName[$f.Name] = $f }

$target = $null
if ($FunctionName) {
    foreach ($f in $fmap) { if ($f.Name -eq $FunctionName) { $target = $f; break } }
    if (-not $target) { throw ("В объекте нет функции {0}. Всего функций: {1}." -f $FunctionName, $fmap.Count) }
}
else {
    foreach ($f in $fmap) { if ($null -eq $target -or $f.CodeLines -gt $target.CodeLines) { $target = $f } }
}

# состояние эмиссии
$script:Listing      = $listing
$script:FuncByName   = $funcByName
$script:CalleeCache  = @{}
$script:Rows         = New-Object System.Collections.Generic.List[string]
$script:Expected     = @{}
$script:Now          = ([datetime]'2026-08-14T10:00:00').Ticks
$script:Rec          = 1000L
$script:StmtIx       = 0
$script:SqlIx        = 0
$script:MainRows     = 0
$script:NoiseIx      = 0
$script:Bias         = $LineNumberBias
$script:BaseTicks    = $BaseTicks
$script:StepTicks    = $StepTicks
$script:CycleLen     = $CycleLen
$script:FrameOverhead= $FrameOverheadTicks
$script:SqlBase      = $SqlBaseTicks
$script:SqlStep      = $SqlStepTicks
$script:MainSession  = $SessionId
$script:NoiseSession = $NoiseSessionId
$script:NoiseEvery   = $NoiseEvery
$script:NoiseOn      = (-not $NoNoise)
$script:ObjTypeS     = $ObjectType
$script:ObjIdS       = $ObjectId
$script:ForeignType  = $ForeignObjectType
$script:ForeignId    = $ForeignObjectId
$script:CalleeLines  = $CalleeLines
$script:MaxDepth     = $MaxCallDepth
$script:TargetFunc   = $target.Name
$script:RecTrigger   = 0
$script:RecStmts     = @()

$script:ForeignStmts = @(
    [pscustomobject]@{ Line = 20; EmitLine = 20; Text = 'CLEAR(Buffer)';  LineCount = 1; IsSql = $false; SqlOp = ''; Callee = ''; CalleeKind = '' }
    [pscustomobject]@{ Line = 21; EmitLine = 21; Text = 'EXIT(TRUE)';     LineCount = 1; IsSql = $false; SqlOp = ''; Callee = ''; CalleeKind = '' }
)

# операторы функции и цикл в ней
$stmts = New-SynthStatementList -Listing $listing -From $target.FirstLine -To $target.LastLine -FuncNames $funcSet
if ($stmts.Count -eq 0) { throw ("В функции {0} нет исполняемых операторов." -f $target.Name) }

$loopStart = -1; $loopEnd = -1
for ($i = 0; $i -lt $stmts.Count; $i++) {
    if ($stmts[$i].Text -match '^REPEAT\b') { $loopStart = $i; break }
}
if ($loopStart -ge 0) {
    $d = 0
    for ($j = $loopStart + 1; $j -lt $stmts.Count; $j++) {
        if     ($stmts[$j].Text -match '^REPEAT\b') { $d++ }
        elseif ($stmts[$j].Text -match '^UNTIL\b')  { if ($d -eq 0) { $loopEnd = $j; break } else { $d-- } }
    }
}
$loopKind = 'REPEAT..UNTIL'
if ($loopStart -lt 0 -or $loopEnd -lt 0) {
    # в функции цикла нет: повторяем срез операторов, чтобы проверка числа вызовов всё равно состоялась
    $loopKind  = 'срез операторов (цикла в функции нет)'
    $loopStart = 0
    $loopEnd   = [math]::Min($stmts.Count - 1, $LoopBodyMax + 1)
}

$body = New-Object System.Collections.Generic.List[object]
for ($i = $loopStart + 1; $i -lt $loopEnd -and $body.Count -lt $LoopBodyMax; $i++) { [void]$body.Add($stmts[$i]) }
if ($body.Count -eq 0) { throw 'Тело цикла пустое — нечего повторять.' }

if ($CompositeProbe -and $body.Count -ge 2) { $body[1].EmitLine = $body[0].Line }

if ($RecursionDepth -gt 0) {
    $script:RecTrigger = $body[0].EmitLine
    $n = [math]::Min(2, $body.Count)
    $rs = @()
    for ($i = 0; $i -lt $n; $i++) { $rs += $body[$i] }
    $script:RecStmts = $rs
}

# программа: до цикла, REPEAT один раз, тело+UNTIL по итерации, после цикла
$program = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $loopStart; $i++)  { [void]$program.Add($stmts[$i]) }
[void]$program.Add($stmts[$loopStart])
for ($it = 0; $it -lt $Iterations; $it++) {
    foreach ($s in $body) { [void]$program.Add($s) }
    [void]$program.Add($stmts[$loopEnd])
}
for ($i = $loopEnd + 1; $i -lt $stmts.Count; $i++) { [void]$program.Add($stmts[$i]) }

# фоновая сессия: другая функция того же объекта, строки которой не пересекаются с телом прогона
$noiseFunc = $null
foreach ($f in ($fmap | Sort-Object CodeLines -Descending)) {
    if ($f.Name -eq $target.Name) { continue }
    if ($f.FirstLine -le $target.LastLine -and $f.LastLine -ge $target.FirstLine) { continue }
    $used = $false
    foreach ($s in $program) { if ($s.Callee -eq $f.Name) { $used = $true; break } }
    if ($used) { continue }
    $noiseFunc = $f; break
}
if ($null -eq $noiseFunc) { $script:NoiseOn = $false }

$script:NoiseSteps = @()
if ($script:NoiseOn) {
    $nl = New-SynthStatementList -Listing $listing -From $noiseFunc.FirstLine -To $noiseFunc.LastLine -FuncNames $null
    $steps = @()
    $steps += @{ Id = 400; Name = 'ALFunctionStart'; Func = $noiseFunc.Name; Line = ''; Text = '' }
    $n = [math]::Min(3, $nl.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $steps += @{ Id = 403; Name = 'ALStatement'; Func = $noiseFunc.Name
                     Line = [string]($nl[$i].Line + $LineNumberBias); Text = $nl[$i].Text }
    }
    $steps += @{ Id = 401; Name = 'ALFunctionStop'; Func = $noiseFunc.Name; Line = ''; Text = '' }
    $script:NoiseSteps = $steps
}

# ---------------------------------------------------------------------------
# прогон
# ---------------------------------------------------------------------------

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$startTicks = $script:Now
$rootDur = Invoke-SynthFrame -Session $SessionId -ObjType $ObjectType -ObjId $ObjectId `
                             -Func $target.Name -Stmts $program.ToArray() -Depth 0 `
                             -RecLeft $RecursionDepth -Thread '4120'

# фоновая сессия дорабатывает свой цикл, чтобы не осталось незакрытых рамок
if ($script:NoiseOn) {
    while (($script:NoiseIx % $script:NoiseSteps.Count) -ne 0) { Invoke-SynthNoise }
}
$sw.Stop()

# ---------------------------------------------------------------------------
# файлы
# ---------------------------------------------------------------------------

$tab = [char]9
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Bom   = New-Object System.Text.UTF8Encoding($true)

$evPath  = Join-Path $OutDir 'events.tsv'
$expPath = Join-Path $OutDir 'expected.tsv'

$header = (@('EventId','EventName','TimeCreatedTicks','SessionId','ObjectType','ObjectId',
             'FunctionName','LineNumber','Statement','Level','RecordId','ThreadId','Raw') -join $tab)
[System.IO.File]::WriteAllText($evPath, ($header + "`r`n" + (($script:Rows -join "`r`n") + "`r`n")), $utf8NoBom)

$expRows = New-Object System.Collections.Generic.List[string]
[void]$expRows.Add((@('LineNo','FunctionName','Hits','TotalTicks','SelfTicks','ChildTicks',
                      'SqlCount','SqlTicks','MinTicks','MaxTicks') -join $tab))
foreach ($k in ($script:Expected.Keys | Sort-Object)) {
    $b = $script:Expected[$k]
    $min = $b.Min; if ($min -eq [int64]::MaxValue) { $min = 0L }
    [void]$expRows.Add((@($b.Line, $b.Func, $b.Hits, $b.Total, $b.Self, $b.Child,
                          $b.SqlCount, $b.SqlTicks, $min, $b.Max) -join $tab))
}
[System.IO.File]::WriteAllText($expPath, (($expRows -join "`r`n") + "`r`n"), $utf8Bom)

$expTotalHits = 0L; $expSelf = 0L
foreach ($k in $script:Expected.Keys) { $expTotalHits += $script:Expected[$k].Hits; $expSelf += $script:Expected[$k].Self }

if (-not $Quiet) {
    Write-Host ''
    Write-Host ('Объект:    {0} {1} {2} (компилирован {3} {4}, строк {5})' -f
        (Get-AlTypeName $ObjectType), $ObjectId, $info.Name, $info.Date, $info.Time, $info.Lines)
    Write-Host ('Функция:   {0} — строки {1}..{2}, операторов {3}' -f
        $target.Name, $target.FirstLine, $target.LastLine, $stmts.Count)
    Write-Host ('Цикл:      {0}, строки {1}..{2}, тело {3} операторов x {4} итераций' -f
        $loopKind, $stmts[$loopStart].Line, $stmts[$loopEnd].Line, $body.Count, $Iterations)
    if ($RecursionDepth -gt 0) { Write-Host ('Рекурсия:  строка {0}, глубина {1}' -f $script:RecTrigger, $RecursionDepth) }
    if ($CompositeProbe)       { Write-Host ('Составная: оператор строки {0} эмитится на строке {1}' -f $body[1].Line, $body[1].EmitLine) }
    if ($script:NoiseOn)       { Write-Host ('Фон:       сессия {0}, функция {1}, событие каждые {2} событий главной' -f $NoiseSessionId, $noiseFunc.Name, $NoiseEvery) }
    Write-Host ('Смещение:  номер строки в событии = строка листинга {0}{1} (разбор должен найти {2}{3})' -f
        $(if ($LineNumberBias -ge 0) { '+' } else { '' }), $LineNumberBias,
        $(if ((-$LineNumberBias) -ge 0) { '+' } else { '' }), (-$LineNumberBias))
    Write-Host ('Событий:   {0:N0}; строк в эталоне {1:N0}; попаданий {2:N0}' -f
        $script:Rows.Count, $script:Expected.Count, $expTotalHits)
    Write-Host ('Время:     рамка {0:N1} мс модельного времени; сгенерировано за {1:N1} с' -f
        ($rootDur / 10000.0), $sw.Elapsed.TotalSeconds)
    Write-Host ('Файлы:     {0}' -f $evPath)
    Write-Host ('           {0}' -f $expPath)
}

# ---------------------------------------------------------------------------
# сверка с эталоном
# ---------------------------------------------------------------------------

if ($Verify) {
    . (Join-Path $PSScriptRoot 'Lib-AlTrace.ps1')

    Write-Host ''
    Write-Host '=== разбор собственного вывода ==='
    $swv = [System.Diagnostics.Stopwatch]::StartNew()
    $events = Import-AlTraceEvents -Path $evPath
    $mine   = @($events | Where-Object { $_.SessionId -eq $SessionId })

    $cal = Resolve-AlLineOffset -Events $mine -Listing $listing -ObjectType $ObjectType -ObjectId $ObjectId
    $lines = Measure-AlLines -Events $events -LineOffset $cal.Offset -Listing $listing `
                             -ObjectType $ObjectType -ObjectId $ObjectId -SessionId $SessionId
    # сводка считается по ПОЛНОМУ набору событий: разрывы RecordId на отфильтрованном
    # наборе ложны — записи чужих сессий из нумерации канала никуда не деваются
    $summary = Get-AlRunSummary -Events $events
    $swv.Stop()

    # эталон с диска
    $exp = @{}
    $expLines = [System.IO.File]::ReadAllLines($expPath, [System.Text.Encoding]::UTF8)
    for ($i = 1; $i -lt $expLines.Length; $i++) {
        if ($expLines[$i].Length -eq 0) { continue }
        $c = $expLines[$i].Split($tab)
        $exp[[int]$c[0]] = [pscustomobject]@{
            LineNo = [int]$c[0]; FunctionName = $c[1]; Hits = [int64]$c[2]
            TotalTicks = [int64]$c[3]; SelfTicks = [int64]$c[4]; ChildTicks = [int64]$c[5]
            SqlCount = [int64]$c[6]; SqlTicks = [int64]$c[7]; MinTicks = [int64]$c[8]; MaxTicks = [int64]$c[9]
        }
    }
    $act = @{}
    foreach ($r in $lines) { $act[[int]$r.LineNo] = $r }

    $onlyExp = @($exp.Keys | Where-Object { -not $act.ContainsKey($_) })
    $onlyAct = @($act.Keys | Where-Object { -not $exp.ContainsKey($_) })
    $common  = @($exp.Keys | Where-Object { $act.ContainsKey($_) } | Sort-Object)

    $metrics = @('Hits','TotalTicks','SelfTicks','ChildTicks','SqlCount','SqlTicks','MinTicks','MaxTicks')
    $res = @{}
    foreach ($m in $metrics) { $res[$m] = @{ Ok = 0; Bad = 0; MaxDelta = 0L; FirstBad = '' } }
    foreach ($k in $common) {
        foreach ($m in $metrics) {
            $d = [int64]$act[$k].$m - [int64]$exp[$k].$m
            if ($d -eq 0) { $res[$m].Ok++ }
            else {
                $res[$m].Bad++
                if ([math]::Abs($d) -gt $res[$m].MaxDelta) { $res[$m].MaxDelta = [math]::Abs($d) }
                if ($res[$m].FirstBad.Length -eq 0) {
                    $res[$m].FirstBad = ('строка {0}: эталон {1}, разбор {2}' -f $k, $exp[$k].$m, $act[$k].$m)
                }
            }
        }
    }

    Write-Host ''
    Write-Host ('Строк: эталон {0}, разбор {1}, общих {2}, только в эталоне {3}, только в разборе {4}' -f
        $exp.Count, $act.Count, $common.Count, $onlyExp.Count, $onlyAct.Count)
    Write-Host ''
    Write-Host ('{0,-12} {1,8} {2,10} {3,14}  {4}' -f 'Метрика', 'Совпало', 'Разошлось', 'Макс. дельта', 'Первое расхождение')
    Write-Host ('{0}' -f ('-' * 96))
    $allOk = $true
    foreach ($m in $metrics) {
        $r = $res[$m]
        if ($r.Bad -gt 0) { $allOk = $false }
        Write-Host ('{0,-12} {1,8} {2,10} {3,14}  {4}' -f $m, $r.Ok, $r.Bad, $r.MaxDelta, $r.FirstBad)
    }
    if ($onlyExp.Count -gt 0) {
        $allOk = $false
        Write-Host ('Только в эталоне: {0}' -f (($onlyExp | Sort-Object | Select-Object -First 10) -join ', ')) -ForegroundColor Yellow
    }
    if ($onlyAct.Count -gt 0) {
        $allOk = $false
        Write-Host ('Только в разборе: {0}' -f (($onlyAct | Sort-Object | Select-Object -First 10) -join ', ')) -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host '=== отдельные проверки ==='

    # 1. цикл
    $bodyLines = @()
    foreach ($s in $body) { if ($bodyLines -notcontains $s.EmitLine) { $bodyLines += $s.EmitLine } }
    $hitOk = 0
    foreach ($l in $bodyLines) { if ($act.ContainsKey($l) -and $act[$l].Hits -eq ($Iterations * (@($body | Where-Object { $_.EmitLine -eq $l }).Count))) { $hitOk++ } }
    Write-Host ('1. Цикл: строк тела {0}; Hits = {1} x число операторов строки у {2} из {0}; строка UNTIL {3} Hits={4}; строка REPEAT {5} Hits={6}' -f
        $bodyLines.Count, $Iterations, $hitOk, $stmts[$loopEnd].Line,
        $(if ($act.ContainsKey($stmts[$loopEnd].Line)) { $act[$stmts[$loopEnd].Line].Hits } else { 0 }),
        $stmts[$loopStart].Line,
        $(if ($act.ContainsKey($stmts[$loopStart].Line)) { $act[$stmts[$loopStart].Line].Hits } else { 0 }))

    # 2. вложенный вызов: Self не задваивается
    $callLines = @($lines | Where-Object { $_.Callees -gt 0 })
    $idOk = 0; $idBad = 0
    foreach ($r in $callLines) {
        if ($r.TotalTicks -eq ($r.SelfTicks + $r.ChildTicks)) { $idOk++ } else { $idBad++ }
    }
    $sample = $null
    if ($callLines.Count -gt 0) { $sample = ($callLines | Sort-Object ChildTicks -Descending)[0] }
    Write-Host ('2. Вложенный вызов: строк с вызовами {0}; Total = Self + дети у {1}, нарушено {2}' -f
        $callLines.Count, $idOk, $idBad)
    if ($sample) {
        Write-Host ('   пример строки {0}: Hits {1}, Total {2}, Self {3}, дети {4}; эталон Self {5} — {6}' -f
            $sample.LineNo, $sample.Hits, $sample.TotalTicks, $sample.SelfTicks, $sample.ChildTicks,
            $exp[$sample.LineNo].SelfTicks,
            $(if ($sample.SelfTicks -eq $exp[$sample.LineNo].SelfTicks) { 'совпало' } else { 'РАЗОШЛОСЬ' }))
    }

    # 3. самокалибровка
    Write-Host ('3. Калибровка: смещение {0}{1}, совпадений {2:N1} % ({3} точных, {4} по префиксу из {5}); Ok={6}' -f
        $(if ($cal.Offset -ge 0) { '+' } else { '' }), $cal.Offset, $cal.MatchPct,
        $cal.ExactCount, $cal.PrefixCount, $cal.Sampled, $cal.Ok)
    foreach ($shift in @(1, 2, 50)) {
        $probe = New-Object System.Collections.Generic.List[object]
        $n = 0
        foreach ($e in $mine) {
            if ($e.Kind -ne 'Stmt') { continue }
            $c = $e.PSObject.Copy()
            $c.LineNumber = $e.LineNumber + $shift
            [void]$probe.Add($c)
            $n++
            if ($n -ge 400) { break }
        }
        $r = Resolve-AlLineOffset -Events $probe -Listing $listing -ObjectType $ObjectType -ObjectId $ObjectId
        Write-Host ('   сдвиг номеров на +{0}: найдено {1}{2}, совпадений {3:N1} %, Ok={4}{5}' -f
            $shift, $(if ($r.Offset -ge 0) { '+' } else { '' }), $r.Offset, $r.MatchPct, $r.Ok,
            $(if ($r.Fallback) { ('; резервный маппинг: ' + $r.Fallback.Count + ' строк') } else { '' }))
    }

    # 4. резервный маппинг вместо нумерации
    $calBad = Resolve-AlLineOffset -Events $mine -Listing $listing -ObjectType $ObjectType `
                                   -ObjectId $ObjectId -Candidates @(-1, 0, 1) -MinMatchPct 101
    $fbLines = Measure-AlLines -Events $events -FallbackMap $calBad.Fallback -LineSource Fallback `
                               -ObjectType $ObjectType -ObjectId $ObjectId -SessionId $SessionId
    $fbOk = 0; $fbBad = 0
    foreach ($r in $fbLines) {
        if ($exp.ContainsKey([int]$r.LineNo) -and $exp[[int]$r.LineNo].Hits -eq $r.Hits) { $fbOk++ } else { $fbBad++ }
    }
    Write-Host ('4. Резервный маппинг (без нумерации): строк {0}, Hits совпал у {1}, разошёлся у {2}; неоднозначных операторов {3}' -f
        $fbLines.Count, $fbOk, $fbBad, $script:AlMeasureStats.Unresolved)

    # 5. изоляция сессий
    $allLines = Measure-AlLines -Events $events -LineOffset $cal.Offset -ObjectType $ObjectType -ObjectId $ObjectId
    $allMap = @{}
    foreach ($r in $allLines) { $allMap[[int]$r.LineNo] = $r }
    $isoBad = 0
    foreach ($k in $common) {
        if (-not $allMap.ContainsKey($k)) { $isoBad++; continue }
        if ($allMap[$k].Hits -ne $act[$k].Hits -or $allMap[$k].SelfTicks -ne $act[$k].SelfTicks) { $isoBad++ }
    }
    Write-Host ('5. Изоляция сессий: разбор без фильтра дал {0} строк (фон добавил {1}); расхождений на общих строках {2}' -f
        $allLines.Count, ($allLines.Count - $lines.Count), $isoBad)

    # 6. рекурсия и составные строки
    $recLines = @($lines | Where-Object { $_.Recursive })
    $recHits = 0L
    foreach ($r in $recLines) { $recHits += $r.RecursiveHits }
    Write-Host ('6. Рекурсия: рамок {0}; строк с флагом {1}; повторных попаданий {2} (Total на них завышен — это ожидаемо)' -f
        $script:AlMeasureStats.RecursiveFrames, $recLines.Count, $recHits)

    $multi = @($lines | Where-Object { $_.Stmts -gt 1 })
    $maxStmts = 0
    foreach ($r in $lines) { if ($r.Stmts -gt $maxStmts) { $maxStmts = $r.Stmts } }
    Write-Host ('7. Составные строки: строк с несколькими текстами операторов {0}, максимум текстов на строке {1}' -f
        $multi.Count, $maxStmts)
    if ($multi.Count -gt 0) {
        $s2 = ($multi | Sort-Object Hits -Descending)[0]
        Write-Host ('   пример строки {0}: текстов {1}, Hits {2}, эталон Hits {3} — {4}' -f
            $s2.LineNo, $s2.Stmts, $s2.Hits, $exp[[int]$s2.LineNo].Hits,
            $(if ($s2.Hits -eq $exp[[int]$s2.LineNo].Hits) { 'совпало' } else { 'РАЗОШЛОСЬ' }))
    }

    # 8. сводка прогона
    Write-Host ('8. Сводка прогона:')
    Write-AlRunSummary -Summary $summary
    Write-Host ('   разбор занял {0:N1} с' -f $swv.Elapsed.TotalSeconds)

    Write-Host ''
    if ($allOk -and $isoBad -eq 0 -and $idBad -eq 0 -and $cal.Offset -eq (-$LineNumberBias) -and $summary.IsValid) {
        Write-Host 'ИТОГ: разбор сошёлся с эталоном по всем метрикам.' -ForegroundColor Green
    }
    else {
        Write-Host 'ИТОГ: есть расхождения — см. выше.' -ForegroundColor Red
    }
}

if ($PassThru) {
    [pscustomobject]@{
        EventsPath   = $evPath
        ExpectedPath = $expPath
        ObjectType   = $ObjectType
        ObjectId     = $ObjectId
        FunctionName = $target.Name
        SessionId    = $SessionId
        Events       = $script:Rows.Count
        Lines        = $script:Expected.Count
        LoopLine     = $stmts[$loopStart].Line
        LoopBody     = $body.Count
        Iterations   = $Iterations
        Offset       = (-$LineNumberBias)
    }
}
