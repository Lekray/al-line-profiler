#Requires -Version 5.1
<#
.SYNOPSIS
    Разбор трассировки C/AL в построчные метрики. Подключается через dot-sourcing.

.DESCRIPTION
    Вход — events.tsv, который пишет Collect-AlTrace.ps1: одна строка на событие ETW,
    колонки EventId, EventName, TimeCreatedTicks, SessionId, ObjectType, ObjectId,
    FunctionName, LineNumber, Statement, Level, RecordId, ThreadId, Raw.
    Сверка «тот ли листинг» — Test-AlListingMatch, по всем строкам, а не по выборке.
    Колонки ищутся ПО ИМЕНАМ заголовка: порядок в файле не фиксируется, как и порядок
    полей payload в манифесте провайдера.

    Модель времени. Событие оператора (роль ALStatement, в NAV 2018 это Id=403) —
    это ТОЧКА, а не интервал: платформа шлёт его перед выполнением оператора и не
    шлёт парного «конец». Поэтому длительность оператора = метка следующего события
    в его рамке вызова. Пары есть только у вызовов функций (ALFunctionStart/Stop)
    и у обращений к SQL (роли <операция>:Start / <операция>:Stop).

    Как считается строка:
      Total (inclusive) — от события оператора до следующего события в его рамке;
      Self              — Total минус сумма длительностей ВЛОЖЕННЫХ РАМОК, начатых
                          внутри этого оператора;
      Sql               — время парных SQL-событий, начатых внутри оператора; из Self
                          НЕ вычитается (запрос выполняется внутри самого оператора,
                          это его собственное время), а лежит отдельной корзиной;
                          чистое время C/AL — SelfNoSqlMs.

    Решённые тонкости потока событий:
      1. Узел вложенного вызова остаётся на строке ВЫЗЫВАЮЩЕГО: между событием
         оператора-вызова и событием входа в функцию своего события нет. Значит Self
         строки вызова после вычитания детей честный, а Total строки вызова — полная
         стоимость вызова вместе с телом вызванной функции.
      2. Рекурсия завышает Total: внешний виток включает время внутреннего, и оба
         ложатся на одни и те же строки. Такие строки помечаются Recursive, число
         повторных попаданий — RecursiveHits. Self при рекурсии честный, завышен
         только Total.
      3. Одна строка может дать НЕСКОЛЬКО событий оператора (составной оператор,
         повторно вычисляемый заголовок цикла, несколько операторов в строке). Всё
         агрегируется суммой, Hits считает события, число разных текстов — Stmts.
      4. Порядок разбора берётся из RecordId (порядок записи в канал), а не из метки
         времени: у соседних событий метки совпадают сплошь и рядом.
      5. Сессии разбираются независимыми стеками: в канале события параллельных
         сессий перемешаны.
      6. Ошибка C/AL — это ВЫХОД из функции, а не пометка на полях: платформа шлёт
         ALFunctionError вместо ALFunctionStop. Рамка снимается по нему так же, как
         по выходу, иначе строка вызова забирает себе весь остаток трассы.
      7. SQL-операция не переживает свой оператор. Если Stop потерян, огрызок
         снимается при закрытии оператора: иначе список открытых операций отравлен
         навсегда и SQL-время сессии перестаёт учитываться — молча, без ошибок.

    Подключение:  . (Join-Path $PSScriptRoot 'Lib-AlTrace.ps1')
#>

$script:AlTicksPerMs   = 10000.0
$script:AlMeasureStats = $null

# ---------------------------------------------------------------------------
# текст
# ---------------------------------------------------------------------------

# Имена типов объектов в payload событий -> номера, как в dbo.[Object].[Type].
# Сравнение ключей в хэш-таблице PowerShell регистронезависимо, поэтому
# «CodeUnit» и «Codeunit» одинаково находятся.
$script:AlObjTypeByName = @{
    'TableData' = 0; 'Table'    = 1; 'Form'      = 2; 'Report' = 3
    'Dataport'  = 4; 'CodeUnit' = 5; 'XMLport'   = 6; 'MenuSuite' = 7
    'Page'      = 8; 'Query'    = 9; 'System'    = 10
}

function ConvertTo-AlNormText {
    <#
    .SYNOPSIS
        Нормализация текста оператора для сравнения: схлопнутые пробелы, без
        хвостового комментария, без завершающих «;».

    .DESCRIPTION
        Комментарий отрезается только ВНЕ строковой константы: '//' встречается
        внутри литералов (адреса, разделители), и наивная обрезка портила бы текст.
        Удвоенная кавычка внутри литерала обрабатывается сама собой: она закрывает
        литерал и тут же открывает следующий.
    #>
    param([string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }

    $sb     = New-Object System.Text.StringBuilder
    $inStr  = $false
    $prevWs = $false
    $n      = $Text.Length

    for ($i = 0; $i -lt $n; $i++) {
        $c = $Text[$i]
        if ($inStr) {
            [void]$sb.Append($c)
            if ($c -eq "'") { $inStr = $false }
            $prevWs = $false
            continue
        }
        if ($c -eq "'") { $inStr = $true; [void]$sb.Append($c); $prevWs = $false; continue }
        if ($c -eq '/' -and ($i + 1) -lt $n -and $Text[$i + 1] -eq '/') { break }
        if ([char]::IsWhiteSpace($c)) {
            if (-not $prevWs) { [void]$sb.Append(' '); $prevWs = $true }
            continue
        }
        [void]$sb.Append($c)
        $prevWs = $false
    }

    $s = $sb.ToString().Trim()
    while ($s.EndsWith(';')) { $s = $s.Substring(0, $s.Length - 1).TrimEnd() }
    return $s
}

function Test-AlTextMatch {
    <#
    .SYNOPSIS
        Сравнение нормализованных текстов: 2 — точное, 1 — по префиксу, 0 — нет.

    .DESCRIPTION
        Префикс нужен потому, что оператор бывает многострочным: платформа шлёт его
        целиком одним текстом, а в строке листинга лежит только первая физическая
        строка. Слишком короткие фрагменты по префиксу не сравниваются — они дали бы
        ложные совпадения на END, EXIT и им подобных.
    #>
    param([string] $Statement, [string] $Line, [int] $MinPrefix = 4)

    if ([string]::IsNullOrEmpty($Statement) -or [string]::IsNullOrEmpty($Line)) { return 0 }
    if ([string]::Equals($Statement, $Line, 'OrdinalIgnoreCase')) { return 2 }
    if ($Statement.Length -gt $Line.Length -and $Line.Length -ge $MinPrefix -and
        $Statement.StartsWith($Line, 'OrdinalIgnoreCase')) { return 1 }
    if ($Line.Length -gt $Statement.Length -and $Statement.Length -ge $MinPrefix -and
        $Line.StartsWith($Statement, 'OrdinalIgnoreCase')) { return 1 }
    return 0
}

# ---------------------------------------------------------------------------
# 1. чтение events.tsv
# ---------------------------------------------------------------------------

function Import-AlTraceEvents {
    <#
    .SYNOPSIS
        Читает events.tsv в набор событий; фильтр по сессии и по объекту.

    .DESCRIPTION
        Колонки берутся по именам из заголовка. Роль события (колонка EventName) уже
        разрешена сборщиком по именам полей манифеста, поэтому номера событий здесь
        не участвуют вовсе; если роль незнакомая, но в событии есть номер строки и
        текст оператора, оно всё равно считается оператором.

        ВНИМАНИЕ: фильтр -ObjectType/-ObjectId здесь — для просмотра. Measure-AlLines
        должен получать поток БЕЗ фильтра по объекту, иначе вырезанные события входа
        и выхода чужих объектов разорвут дерево вызовов и время вложенных вызовов
        уедет в Self вызывающей строки. Свой фильтр по объекту у Measure-AlLines есть,
        он применяется уже к результату.

    .PARAMETER Path
        Путь к events.tsv.

    .PARAMETER SessionId
        Оставить только эти сессии (значения колонки SessionId как строки).

    .PARAMETER ObjectType
        Оставить события этого типа объекта; события без объекта (SQL, служебные)
        сохраняются всегда.

    .PARAMETER ObjectId
        Оставить события этого номера объекта.

    .PARAMETER MaxEvents
        Прочитать не больше N строк (0 — все).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Path,
        [string[]] $SessionId,
        [int]      $ObjectType = -1,
        [int]      $ObjectId   = -1,
        [int]      $MaxEvents  = 0
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw ('Нет файла трассировки: ' + $Path) }

    $sessSet = $null
    if ($SessionId -and @($SessionId).Count -gt 0) {
        $sessSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($s in $SessionId) { [void]$sessSet.Add(([string]$s).Trim()) }
    }

    $list = New-Object System.Collections.Generic.List[object]
    $sr   = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
    try {
        $head = $sr.ReadLine()
        if ($null -eq $head) { throw ('Файл трассировки пуст: ' + $Path) }
        $cols = $head.Split([char]9)
        $ix   = @{}
        for ($i = 0; $i -lt $cols.Length; $i++) { $ix[$cols[$i].Trim()] = $i }
        foreach ($req in @('EventId', 'EventName', 'TimeCreatedTicks')) {
            if (-not $ix.ContainsKey($req)) {
                throw ('В заголовке events.tsv нет колонки {0}; заголовок: {1}' -f $req, $head)
            }
        }
        $iId = $ix['EventId']; $iNm = $ix['EventName']; $iTk = $ix['TimeCreatedTicks']
        $iSs = -1; if ($ix.ContainsKey('SessionId'))    { $iSs = $ix['SessionId'] }
        $iOt = -1; if ($ix.ContainsKey('ObjectType'))   { $iOt = $ix['ObjectType'] }
        $iOi = -1; if ($ix.ContainsKey('ObjectId'))     { $iOi = $ix['ObjectId'] }
        $iFn = -1; if ($ix.ContainsKey('FunctionName')) { $iFn = $ix['FunctionName'] }
        $iLn = -1; if ($ix.ContainsKey('LineNumber'))   { $iLn = $ix['LineNumber'] }
        $iSt = -1; if ($ix.ContainsKey('Statement'))    { $iSt = $ix['Statement'] }
        $iRi = -1; if ($ix.ContainsKey('RecordId'))     { $iRi = $ix['RecordId'] }
        $iTh = -1; if ($ix.ContainsKey('ThreadId'))     { $iTh = $ix['ThreadId'] }
        $iRw = -1; if ($ix.ContainsKey('Raw'))          { $iRw = $ix['Raw'] }

        $rxSql = New-Object System.Text.RegularExpressions.Regex('(?:^|;\s*)sqlStatement=', 'Compiled')

        $seq = 0
        $ln  = $null
        while ($null -ne ($ln = $sr.ReadLine())) {
            if ($ln.Length -eq 0) { continue }
            $f  = $ln.Split([char]9)
            $fc = $f.Length
            $seq++
            if ($MaxEvents -gt 0 -and $seq -gt $MaxEvents) { break }

            $vSes = ''; if ($iSs -ge 0 -and $iSs -lt $fc) { $vSes = $f[$iSs] }
            if ($sessSet -and -not $sessSet.Contains($vSes)) { continue }

            $vNm = ''; if ($iNm -lt $fc) { $vNm = $f[$iNm] }
            $vFn = ''; if ($iFn -ge 0 -and $iFn -lt $fc) { $vFn = $f[$iFn] }
            $vSt = ''; if ($iSt -ge 0 -and $iSt -lt $fc) { $vSt = $f[$iSt] }
            $vRw = ''; if ($iRw -ge 0 -and $iRw -lt $fc) { $vRw = $f[$iRw] }
            $vTh = ''; if ($iTh -ge 0 -and $iTh -lt $fc) { $vTh = $f[$iTh] }

            $nId = 0;  [void][int]::TryParse($f[$iId], [ref]$nId)
            $nTk = 0L; [void][int64]::TryParse($f[$iTk], [ref]$nTk)
            $nRi = 0L; if ($iRi -ge 0 -and $iRi -lt $fc) { [void][int64]::TryParse($f[$iRi], [ref]$nRi) }
            $nOt = -1
            if ($iOt -ge 0 -and $iOt -lt $fc -and $f[$iOt].Length -gt 0) {
                if (-not [int]::TryParse($f[$iOt], [ref]$nOt)) {
                    # В ЖИВЫХ событиях objectType приходит ИМЕНЕМ («Report», «CodeUnit»),
                    # а в синтетике — числом. Без приведения фильтр по объекту молча
                    # не срабатывает, и метрик получается ноль при исправном трейсе.
                    if ($script:AlObjTypeByName.ContainsKey($f[$iOt])) {
                        $nOt = $script:AlObjTypeByName[$f[$iOt]]
                    } else { $nOt = -1 }
                }
            }
            $nOi = -1; if ($iOi -ge 0 -and $iOi -lt $fc -and $f[$iOi].Length -gt 0) { [void][int]::TryParse($f[$iOi], [ref]$nOi) }
            $hasLine = ($iLn -ge 0 -and $iLn -lt $fc -and $f[$iLn].Length -gt 0)
            $nLn = -1; if ($hasLine) { [void][int]::TryParse($f[$iLn], [ref]$nLn) }

            if ($ObjectType -ge 0 -and $nOt -ge 0 -and $nOt -ne $ObjectType) { continue }
            if ($ObjectId   -ge 0 -and $nOi -ge 0 -and $nOi -ne $ObjectId)   { continue }

            # роль -> класс события; номер события не участвует
            $op = ''
            if     ($vNm -eq 'ALStatement')          { $kind = 'Stmt' }
            elseif ($vNm -eq 'ALFunctionStart')      { $kind = 'Start' }
            elseif ($vNm -eq 'ALFunctionStop')       { $kind = 'Stop' }
            elseif ($vNm -eq 'ALFunctionError')      { $kind = 'Error' }
            elseif ($vNm.EndsWith(':Start'))         { $kind = 'PairStart'; $op = $vNm.Substring(0, $vNm.Length - 6) }
            elseif ($vNm.EndsWith(':Stop'))          { $kind = 'PairStop';  $op = $vNm.Substring(0, $vNm.Length - 5) }
            elseif ($hasLine -and $vSt.Length -gt 0) { $kind = 'Stmt' }
            else                                     { $kind = 'Other' }
            if ($op.StartsWith('Sql:')) { $op = $op.Substring(4) }

            $sql = ''
            if (($kind -eq 'PairStart' -or $kind -eq 'PairStop') -and $vRw.Length -gt 0) {
                # sqlStatement — последнее сохраняемое поле payload у SQL-событий,
                # поэтому берётся весь хвост: сам текст запроса содержит «; »
                $m = $rxSql.Match($vRw)
                if ($m.Success) { $sql = $vRw.Substring($m.Index + $m.Length) }
            }

            [void]$list.Add([pscustomobject]@{
                Seq          = $seq
                EventId      = $nId
                Role         = $vNm
                Kind         = $kind
                OpName       = $op
                Ticks        = $nTk
                SessionId    = $vSes
                ObjectType   = $nOt
                ObjectId     = $nOi
                FunctionName = $vFn
                LineNumber   = $nLn
                Statement    = $vSt
                RecordId     = $nRi
                ThreadId     = $vTh
                SqlText      = $sql
            })
        }
    }
    finally { $sr.Dispose() }

    return $list
}

# ---------------------------------------------------------------------------
# 2. самокалибровка нумерации
# ---------------------------------------------------------------------------

function Resolve-AlLineOffset {
    <#
    .SYNOPSIS
        Подбирает смещение «строка листинга = lineNumber события + смещение»
        сличением текста оператора с текстом строки.

    .DESCRIPTION
        Формула +1 выведена по коду ядра (Code Coverage хранит Line No. = lineNumber + 1),
        но эмпирически не подтверждена, а на другой версии платформы может отличаться.
        Поэтому смещение подбирается на каждом прогоне: берутся первые -SampleSize
        событий оператора, для каждого кандидата нормализованный текст оператора
        сличается с нормализованным текстом строки листинга; побеждает кандидат с
        наибольшей долей совпадений.

        Ничья. Прежде она решалась в пользу МЕНЬШЕГО |смещения|, и это молчаливый
        отказ: если вся выборка встала на строки-дубли (в листинге они есть, порядка
        полутора процентов), кандидаты совпадают дословно, и ноль обгоняет верную
        единицу - неверное смещение уходит в отчёт как стопроцентно достоверное.
        Меньшее |смещение| - не довод, а привычка. Ничья означает ровно одно: в
        выборке нет РАЗЛИЧАЮЩИХ событий, то есть таких, у которых текст строки при
        разных смещениях разный. Они ищутся по всей трассе, а не по первым
        -SampleSize событиям, и решают спор текстом. Если различающих событий нет
        вовсе - Ok=$false, Ambiguous=$true, и в дело идёт резервный маппинг.

        Доля ниже -MinMatchPct — Ok=$false. Тогда возвращается резервный маппинг
        Fallback: хеш-таблица «имя функции + TAB + ВЕРХНИЙ_РЕГИСТР(нормализованный
        текст) -> номер строки», построенная только по ОДНОЗНАЧНЫМ парам (текст,
        встречающийся в функции дважды, отброшен). Он передаётся в Measure-AlLines
        параметром -FallbackMap и работает без нумерации вовсе.

    .PARAMETER Events
        Набор событий из Import-AlTraceEvents.

    .PARAMETER Listing
        Листинг объекта из Get-AlListing (Lib-AlListing.ps1).

    .PARAMETER Candidates
        Проверяемые смещения.

    .PARAMETER SampleSize
        Сколько первых событий оператора взять в выборку.

    .PARAMETER MinMatchPct
        Порог доли совпадений, ниже которого смещение считается неподобранным.

    .PARAMETER ObjectType
        Тип объекта, по которому калиброваться.

    .PARAMETER ObjectId
        Номер объекта, по которому калиброваться.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]] $Events,
        [Parameter(Mandatory)][object[]] $Listing,
        [int[]]  $Candidates  = @(-1, 0, 1),
        [int]    $SampleSize  = 200,
        [double] $MinMatchPct = 80.0,
        [int]    $ObjectType  = -1,
        [int]    $ObjectId    = -1
    )

    # текст листинга по номеру строки, нормализуется один раз
    $maxLine = 0
    foreach ($l in $Listing) { if ($l.LineNo -gt $maxLine) { $maxLine = $l.LineNo } }
    $norm = New-Object 'string[]' ($maxLine + 2)
    foreach ($l in $Listing) { $norm[$l.LineNo] = ConvertTo-AlNormText $l.Text }

    $sample = New-Object System.Collections.Generic.List[object]
    foreach ($e in $Events) {
        if ($e.Kind -ne 'Stmt') { continue }
        if ($ObjectType -ge 0 -and $e.ObjectType -ne $ObjectType) { continue }
        if ($ObjectId   -ge 0 -and $e.ObjectId   -ne $ObjectId)   { continue }
        if ($e.LineNumber -lt 0) { continue }
        [void]$sample.Add($e)
        if ($sample.Count -ge $SampleSize) { break }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $best = $null
    foreach ($c in $Candidates) {
        $exact = 0; $prefix = 0
        foreach ($e in $sample) {
            $ln = $e.LineNumber + $c
            if ($ln -lt 1 -or $ln -gt $maxLine) { continue }
            $r = Test-AlTextMatch (ConvertTo-AlNormText $e.Statement) $norm[$ln]
            if     ($r -eq 2) { $exact++ }
            elseif ($r -eq 1) { $prefix++ }
        }
        $pct = 0.0
        if ($sample.Count -gt 0) { $pct = 100.0 * ($exact + $prefix) / $sample.Count }
        $row = [pscustomobject]@{ Offset = $c; MatchPct = $pct; ExactCount = $exact
                                  PrefixCount = $prefix; Decisive = 0 }
        [void]$rows.Add($row)
        # при равной доле совпадений побеждает большее число ТОЧНЫХ; дальше - ничья,
        # и разбирается она ниже текстом, а не величиной смещения
        if ($null -eq $best) { $best = $row }
        elseif ($row.MatchPct -gt $best.MatchPct) { $best = $row }
        elseif ($row.MatchPct -eq $best.MatchPct -and $row.ExactCount -gt $best.ExactCount) { $best = $row }
    }

    # Разбор ничьей. Кандидаты, совпавшие и по доле, и по числу точных, неразличимы
    # на этой выборке - значит выборка состоит из строк, которые при разных смещениях
    # выглядят одинаково. Ищем событие, у которого тексты строк при этих смещениях
    # РАЗНЫЕ: только оно и может рассудить. Идём по всей трассе - в выборку такие
    # события не попали именно потому, что она обрезана первыми -SampleSize.
    $tie = @()
    foreach ($r in $rows) {
        if ($r.MatchPct -eq $best.MatchPct -and $r.ExactCount -eq $best.ExactCount) { $tie += $r }
    }
    $ambiguous = $false

    if ($tie.Count -gt 1 -and $best.MatchPct -gt 0) {
        $decided = 0
        foreach ($e in $Events) {
            if ($decided -ge $SampleSize) { break }
            if ($e.Kind -ne 'Stmt') { continue }
            if ($ObjectType -ge 0 -and $e.ObjectType -ne $ObjectType) { continue }
            if ($ObjectId   -ge 0 -and $e.ObjectId   -ne $ObjectId)   { continue }
            if ($e.LineNumber -lt 0) { continue }

            $t0 = $null; $differs = $false; $firstCand = $true
            foreach ($r in $tie) {
                $ln = $e.LineNumber + $r.Offset
                $tx = ''
                if ($ln -ge 1 -and $ln -le $maxLine) { $tx = [string]$norm[$ln] }
                if ($firstCand) { $t0 = $tx; $firstCand = $false }
                elseif ($tx -ne $t0) { $differs = $true }
            }
            if (-not $differs) { continue }

            $decided++
            $st = ConvertTo-AlNormText $e.Statement
            foreach ($r in $tie) {
                $ln = $e.LineNumber + $r.Offset
                if ($ln -lt 1 -or $ln -gt $maxLine) { continue }
                if ((Test-AlTextMatch $st ([string]$norm[$ln])) -gt 0) { $r.Decisive = $r.Decisive + 1 }
            }
        }

        $win = $null; $draw = $false
        foreach ($r in $tie) {
            if ($null -eq $win -or $r.Decisive -gt $win.Decisive) { $win = $r; $draw = $false }
            elseif ($r.Decisive -eq $win.Decisive) { $draw = $true }
        }
        if ($null -ne $win -and $win.Decisive -gt 0 -and -not $draw) { $best = $win }
        else { $ambiguous = $true }
    }

    $ok       = ($sample.Count -gt 0 -and $null -ne $best -and
                 $best.MatchPct -ge $MinMatchPct -and -not $ambiguous)
    $msg      = ''
    $fallback = $null

    if (-not $ok) {
        $seen = @{}
        foreach ($l in $Listing) {
            if ($l.Kind -ne 'Code') { continue }
            $t = $norm[$l.LineNo]
            if ($t.Length -eq 0) { continue }
            $k = $l.FunctionName + "`t" + $t.ToUpperInvariant()
            if ($seen.ContainsKey($k)) { $seen[$k] = -1 } else { $seen[$k] = $l.LineNo }
        }
        $fallback = @{}
        foreach ($k in $seen.Keys) { if ($seen[$k] -gt 0) { $fallback[$k] = $seen[$k] } }

        if ($sample.Count -eq 0) {
            $msg = 'В трассировке нет событий оператора по этому объекту — калибровать не на чем.'
        }
        elseif ($ambiguous) {
            $lst = @()
            foreach ($r in $tie) { $lst += ('{0}{1}' -f $(if ($r.Offset -gt 0) { '+' } else { '' }), $r.Offset) }
            $msg = ('Смещения {0} совпали одинаково ({1:N1} %), и ни одного различающего события ' +
                    'в трассе нет: строки объекта в этих местах повторяются дословно. Выбирать ' +
                    'смещение не на чем, меньшее по модулю — не довод. Резервный маппинг ' +
                    '(функция + текст) построен, однозначных строк {2}; передайте его в ' +
                    'Measure-AlLines: -FallbackMap <карта> -LineSource Fallback.') -f
                    ($lst -join ' и '), $best.MatchPct, $fallback.Count
        }
        else {
            $msg = ('Совпадений {0:N1} % при пороге {1:N0} %. Нумерация не сошлась: объект в базе ' +
                    'пересобран после дампа либо версия платформы другая. Резервный маппинг ' +
                    '(функция + текст) построен, однозначных строк {2}; передайте его в ' +
                    'Measure-AlLines: -FallbackMap <карта> -LineSource Fallback.') -f
                    $best.MatchPct, $MinMatchPct, $fallback.Count
        }
    }

    $bOff = 1; $bPct = 0.0; $bEx = 0; $bPre = 0
    if ($null -ne $best) { $bOff = $best.Offset; $bPct = $best.MatchPct; $bEx = $best.ExactCount; $bPre = $best.PrefixCount }

    return [pscustomobject]@{
        Offset      = $bOff
        MatchPct    = [math]::Round($bPct, 2)
        ExactCount  = $bEx
        PrefixCount = $bPre
        Sampled     = $sample.Count
        Ok          = $ok
        Ambiguous   = $ambiguous
        Candidates  = $rows
        Fallback    = $fallback
        Message     = $msg
    }
}

function Test-AlListingMatch {
    <#
    .SYNOPSIS
        Сверка листинга с трассой по ВСЕМ строкам объекта, а не по выборке
        калибровки. Отвечает на вопрос «тот ли это листинг», а не «какое смещение».

    .DESCRIPTION
        Дамп исходников и трасса могут разъехаться: объект перекомпилировали между
        дампом и сценарием, отчёт пересобрали по вчерашнему каталогу прогона (общий
        .alsrc с тех пор переписан), сбор запустили с -SkipDump поверх старого дампа.
        Номера строк тогда врут, а выглядит всё нормально.

        Признак, на который в шапке Rebuild-Report ссылались раньше, — «процент
        совпадений при калибровке» — эту проверку НЕ делает: калибровка берёт первые
        -SampleSize событий. На реальном прогоне первые 200 событий покрывали строки
        объекта примерно до 501-й, тогда как события доходили до 546-й: правка ниже
        501-й строки калибровкой не видна вовсе. И модель у калибровки не та: одна
        константа из {-1,0,1} на весь файл вставку строки не описывает.

        Здесь сличаются все РАЗЛИЧНЫЕ пары «номер строки + текст оператора» по всей
        трассе. Единица учёта — СТРОКА, а не оператор: в одной строке C/AL бывает
        несколько операторов ('IF Cond THEN Foo;'), платформа шлёт их порознь, и текст
        второго со строкой листинга не совпадёт никогда. Строка засчитана, если лёг
        хотя бы один её оператор. Сдвиг нумерации так виден, а многооператорная
        строка ложной тревоги не даёт.

        Hash из index.tsv ([Object Metadata].[Hash] платформы) здесь не участвует:
        сверять его в момент сборки отчёта не с чем — метаданных объекта на момент
        сбора в каталоге прогона нет. Отметку кладёт Rebuild-Report, и она ловит
        подмену дампа между сборками; эта же проверка ловит расхождение по существу.

    .PARAMETER LineOffset
        Смещение, выбранное калибровкой: строка листинга = lineNumber + смещение.

    .PARAMETER MinMatchPct
        Доля совпавших строк, ниже которой листинг признаётся чужим.

    .PARAMETER SampleCount
        Сколько несовпавших строк вернуть для диагноза.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]] $Events,
        [Parameter(Mandatory)][object[]] $Listing,
        [int]    $LineOffset  = 1,
        [int]    $ObjectType  = -1,
        [int]    $ObjectId    = -1,
        [double] $MinMatchPct = 80.0,
        [int]    $SampleCount = 5
    )

    $maxLine = 0
    foreach ($l in $Listing) { if ($l.LineNo -gt $maxLine) { $maxLine = $l.LineNo } }
    $norm = New-Object 'string[]' ($maxLine + 2)
    foreach ($l in $Listing) { $norm[$l.LineNo] = ConvertTo-AlNormText $l.Text }

    $seen     = @{}   # пара «строка + текст» уже сверена
    $lineOk   = @{}   # строка листинга -> лёг ли на неё хоть один оператор
    $lineBad  = @{}   # строка листинга -> первый не совпавший текст, для диагноза

    foreach ($e in $Events) {
        if ($e.Kind -ne 'Stmt') { continue }
        if ($ObjectType -ge 0 -and $e.ObjectType -ne $ObjectType) { continue }
        if ($ObjectId   -ge 0 -and $e.ObjectId   -ne $ObjectId)   { continue }
        if ($e.LineNumber -lt 0) { continue }

        $k = '{0}|{1}' -f $e.LineNumber, $e.Statement
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true

        $ln  = $e.LineNumber + $LineOffset
        $hit = $false
        if ($ln -ge 1 -and $ln -le $maxLine) {
            $hit = ((Test-AlTextMatch (ConvertTo-AlNormText $e.Statement) ([string]$norm[$ln])) -gt 0)
        }
        if ($hit) { $lineOk[$ln] = $true }
        else {
            if (-not $lineOk.ContainsKey($ln)) { $lineOk[$ln] = $false }
            if (-not $lineBad.ContainsKey($ln)) { $lineBad[$ln] = $e.Statement }
        }
    }

    $checked  = $lineOk.Count
    $matched  = 0
    $firstBad = 0
    $samples  = New-Object System.Collections.Generic.List[object]
    foreach ($ln in @($lineOk.Keys | Sort-Object)) {
        if ($lineOk[$ln]) { $matched++; continue }
        if ($firstBad -eq 0) { $firstBad = $ln }
        if ($samples.Count -lt $SampleCount) {
            $txt = ''
            if ($ln -ge 1 -and $ln -le $maxLine) { $txt = [string]$norm[$ln] }
            [void]$samples.Add([pscustomobject]@{
                LineNo = $ln; Statement = $lineBad[$ln]; Listing = $txt })
        }
    }

    $pct = 0.0
    if ($checked -gt 0) { $pct = 100.0 * $matched / $checked }
    # сверять нечего — это не расхождение: облегчённый сбор операторов не шлёт вовсе
    $ok = ($checked -eq 0 -or $pct -ge $MinMatchPct)

    return [pscustomobject]@{
        Checked     = $checked
        Matched     = $matched
        MatchPct    = [math]::Round($pct, 2)
        FirstBad    = $firstBad
        LineOffset  = $LineOffset
        Ok          = $ok
        Samples     = $samples.ToArray()
    }
}

# ---------------------------------------------------------------------------
# 3. построчные метрики
# ---------------------------------------------------------------------------

function New-AlLineBucket {
    <#
    .SYNOPSIS
        Пустая копилка метрик одной строки кода.
    #>
    param([int] $ObjectType, [int] $ObjectId, [int] $LineNo)
    return @{
        ObjectType = $ObjectType; ObjectId = $ObjectId; LineNo = $LineNo
        FunctionName = ''
        Hits = 0L; TotalTicks = 0L; SelfTicks = 0L; ChildTicks = 0L
        MinTicks = [int64]::MaxValue; MaxTicks = 0L
        SqlCount = 0L; SqlTicks = 0L; Callees = 0L
        Recursive = $false; RecursiveHits = 0L; NegativeSelf = 0L
        Texts = (New-Object 'System.Collections.Generic.HashSet[string]')
    }
}

function Get-AlLineBucket {
    <#
    .SYNOPSIS
        Копилка строки по ключу «тип/номер объекта/строка», создаётся при первом обращении.
    #>
    param([hashtable] $Buckets, [int] $ObjectType, [int] $ObjectId, [int] $LineNo)
    $k = '{0}/{1}/{2}' -f $ObjectType, $ObjectId, $LineNo
    $b = $Buckets[$k]
    if ($null -eq $b) {
        $b = New-AlLineBucket -ObjectType $ObjectType -ObjectId $ObjectId -LineNo $LineNo
        $Buckets[$k] = $b
    }
    return $b
}

function Close-AlFrameStatement {
    <#
    .SYNOPSIS
        Закрывает открытый оператор рамки на метке $EndTicks и раскладывает его по строке.
    #>
    param([hashtable] $Frame, [int64] $EndTicks, [hashtable] $State, [hashtable] $Stats)

    # Незакрытые SQL-операции этой рамки снимаются ЗДЕСЬ, до выхода по HasStmt.
    # Парного Stop у них уже не будет: оператор, внутри которого они начались,
    # закончился. Оставить их нельзя - список открытых операций отравляется навсегда,
    # позиция следующей операции никогда больше не будет нулевой, и SQL-время сессии
    # перестаёт учитываться совсем, молча и без единого предупреждения. Обратный отказ
    # не лучше: поздний Stop той же операции спарился бы с древним огрызком и записал
    # на его строку фантомные секунды.
    $q = $State.Sql
    for ($i = $q.Count - 1; $i -ge 0; $i--) {
        if ([object]::ReferenceEquals($q[$i].Frame, $Frame)) { $q.RemoveAt($i); $Stats.SqlDropped++ }
    }

    if (-not $Frame.HasStmt) { return }
    $Frame.HasStmt = $false

    $b = $Frame.StmtBucket
    if ($null -eq $b) { return }

    $incl = $EndTicks - $Frame.StmtStart
    if ($incl -lt 0) { $incl = 0L }
    $self = $incl - $Frame.StmtChild
    if ($self -lt 0) { $self = 0L; $b.NegativeSelf++; $Stats.NegativeSelf++ }

    $b.Hits++
    $b.TotalTicks += $incl
    $b.SelfTicks  += $self
    $b.ChildTicks += $Frame.StmtChild
    $b.Callees    += $Frame.StmtCallees
    if ($incl -lt $b.MinTicks) { $b.MinTicks = $incl }
    if ($incl -gt $b.MaxTicks) { $b.MaxTicks = $incl }
    if ($Frame.Recursive) { $b.Recursive = $true; $b.RecursiveHits++ }
    if ($b.FunctionName.Length -eq 0) { $b.FunctionName = $Frame.FunctionName }
    if ($b.Texts.Count -lt 32 -and $Frame.StmtText.Length -gt 0) { [void]$b.Texts.Add($Frame.StmtText) }

    $State.SelfTicks += $self
}

function Pop-AlFrame {
    <#
    .SYNOPSIS
        Снимает верхнюю рамку стека: её длительность уходит в детей открытого оператора
        вызывающего, а у корневой — в измеренное время сессии.
    #>
    param([hashtable] $State, [int64] $EndTicks, [hashtable] $Stats)

    $stack = $State.Stack
    if ($stack.Count -eq 0) { return }
    $fr = $stack[$stack.Count - 1]
    $stack.RemoveAt($stack.Count - 1)

    Close-AlFrameStatement -Frame $fr -EndTicks $EndTicks -State $State -Stats $Stats

    $dur = $EndTicks - $fr.StartTicks
    if ($dur -lt 0) { $dur = 0L }

    if ($stack.Count -gt 0) {
        $p = $stack[$stack.Count - 1]
        if ($p.HasStmt) { $p.StmtChild += $dur; $p.StmtCallees++ }
        else { $State.OrphanChildTicks += $dur }
    }
    else { $State.RootTicks += $dur }
}

function Push-AlFrame {
    <#
    .SYNOPSIS
        Кладёт на стек рамку вызова; помечает рекурсию, если такая же уже в стеке.
    #>
    param([hashtable] $State, $Ev, [int64] $StartTicks, [bool] $Synthetic, [hashtable] $Stats)

    $stack = $State.Stack
    $rec = $false
    for ($i = 0; $i -lt $stack.Count; $i++) {
        if ($stack[$i].ObjectType   -eq $Ev.ObjectType -and
            $stack[$i].ObjectId     -eq $Ev.ObjectId   -and
            $stack[$i].FunctionName -eq $Ev.FunctionName) { $rec = $true; break }
    }
    if ($rec) { $Stats.RecursiveFrames++ }

    $fr = @{
        ObjectType = $Ev.ObjectType; ObjectId = $Ev.ObjectId
        FunctionName = $Ev.FunctionName; StartTicks = $StartTicks
        Recursive = $rec; Synthetic = $Synthetic
        HasStmt = $false; StmtBucket = $null; StmtStart = 0L
        StmtChild = 0L; StmtCallees = 0; StmtText = ''
    }
    [void]$stack.Add($fr)
    if ($stack.Count -gt $Stats.MaxDepth) { $Stats.MaxDepth = $stack.Count }
    return $fr
}

function Measure-AlLines {
    <#
    .SYNOPSIS
        Главная функция: строит дерево вызовов по событиям входа и выхода и считает
        по каждой строке Hits, Total, Self, Min, Max, Avg, SqlCount, SqlMs.

    .DESCRIPTION
        Возвращает по объекту на каждую затронутую строку кода. Диагностика прогона
        (потери, рекурсия, незакрытые рамки) складывается в $script:AlMeasureStats,
        оттуда её берёт Get-AlRunSummary.

        Поток событий разбирается стеком на каждую сессию:
          вход в функцию  — новая рамка; открытый оператор вызывающего НЕ закрывается,
                            он и есть строка вызова;
          оператор        — закрыть предыдущий оператор рамки (его Total = метка этого
                            события минус его собственная), открыть новый;
          выход           — закрыть последний оператор рамки на метке выхода, снять
                            рамку, её длительность записать в детей открытого оператора
                            вызывающего;
          пара Start/Stop — SQL и прочие парные операции: длительность идёт в корзину
                            SqlTicks той строки, на которой операция НАЧАЛАСЬ.

        Дырки в потоке (потерянные события) разбор не роняют: оператор чужой функции
        либо сворачивает стек до своей рамки, либо создаёт синтетическую; выход без
        входа пропускается; всё это считается в диагностике.

    .PARAMETER Events
        Набор событий из Import-AlTraceEvents, БЕЗ фильтра по объекту.

    .PARAMETER LineOffset
        Смещение из Resolve-AlLineOffset: строка листинга = lineNumber + смещение.

    .PARAMETER FallbackMap
        Резервный маппинг «функция + TAB + ВЕРХНИЙ_РЕГИСТР(нормализованный текст) -> строка».

    .PARAMETER LineSource
        Offset — только смещение; Fallback — только резервный маппинг (что не нашлось,
        не учитывается); FallbackThenOffset — сначала маппинг, затем смещение.

    .PARAMETER Listing
        Листинг объекта: если передан, в результат подставляется текст строки.

    .PARAMETER ObjectType
        Фильтр РЕЗУЛЬТАТА по типу объекта (дерево строится по всем событиям).

    .PARAMETER ObjectId
        Фильтр РЕЗУЛЬТАТА по номеру объекта.

    .PARAMETER SessionId
        Разбирать только эти сессии.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]] $Events,
        [int]        $LineOffset = 1,
        [hashtable]  $FallbackMap,
        [ValidateSet('Offset', 'Fallback', 'FallbackThenOffset')]
        [string]     $LineSource = 'Offset',
        [object[]]   $Listing,
        [int]        $ObjectType = -1,
        [int]        $ObjectId   = -1,
        [string[]]   $SessionId
    )

    if ($PSBoundParameters.ContainsKey('FallbackMap') -and
        -not $PSBoundParameters.ContainsKey('LineSource')) { $LineSource = 'FallbackThenOffset' }

    # Индекс резервного маппинга по функциям. Он нужен из-за многострочных операторов:
    # платформа шлёт такой оператор одним текстом, а в маппинге лежит текст ПЕРВОЙ
    # строки, поэтому точного совпадения не будет — ищем самую длинную строку функции,
    # которая является началом текста оператора. Результат запоминается: разных текстов
    # операторов сотни, а событий миллионы.
    $fbIndex = $null
    $fbMemo  = $null
    if ($LineSource -ne 'Offset' -and $FallbackMap) {
        $fbIndex = @{}
        foreach ($k in $FallbackMap.Keys) {
            $p = $k.Split([char]9)
            if ($p.Length -lt 2) { continue }
            if (-not $fbIndex.ContainsKey($p[0])) {
                $fbIndex[$p[0]] = New-Object System.Collections.Generic.List[object]
            }
            [void]$fbIndex[$p[0]].Add([pscustomobject]@{ Text = $p[1]; LineNo = [int]$FallbackMap[$k] })
        }
        $fbMemo = @{}
    }

    $sessSet = $null
    if ($SessionId -and @($SessionId).Count -gt 0) {
        $sessSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($s in $SessionId) { [void]$sessSet.Add(([string]$s).Trim()) }
    }

    $stats = @{
        Events = 0; Stmt = 0; Start = 0; Stop = 0; Error = 0; SqlStart = 0; SqlStop = 0; Other = 0
        OrphanFrames = 0; MissedStarts = 0; UnmatchedStops = 0; ForcedPops = 0
        UnclosedFrames = 0; RecursiveFrames = 0; NegativeSelf = 0; Unresolved = 0
        SqlOrphan = 0; SqlUnmatched = 0; SqlNested = 0; SqlDropped = 0; MaxDepth = 0
        ErrorPops = 0; ErrorUnmatched = 0
        SqlTicks = 0L; SqlCount = 0L; RootTicks = 0L; SelfTicks = 0L
        Lines = 0; AllLines = 0
        Sessions = @{}
    }

    # порядок разбора — порядок записи в канал; сортируем, только если он нарушен
    $need = $false
    $prev = [int64]::MinValue
    foreach ($e in $Events) {
        if ($e.RecordId -lt $prev) { $need = $true; break }
        $prev = $e.RecordId
    }
    if ($need) { $ordered = @($Events | Sort-Object -Property RecordId, Seq) } else { $ordered = $Events }

    $buckets = @{}
    $state   = @{}

    foreach ($ev in $ordered) {
        if ($sessSet -and -not $sessSet.Contains($ev.SessionId)) { continue }
        $stats.Events++

        $st = $state[$ev.SessionId]
        if ($null -eq $st) {
            $st = @{
                Stack = (New-Object System.Collections.Generic.List[object])
                Sql   = (New-Object System.Collections.Generic.List[object])
                First = $ev.Ticks; Last = $ev.Ticks
                RootTicks = 0L; SelfTicks = 0L; OrphanChildTicks = 0L; Events = 0
            }
            $state[$ev.SessionId] = $st
        }
        if ($ev.Ticks -gt $st.Last) { $st.Last = $ev.Ticks }
        $st.Events++

        switch ($ev.Kind) {

            'Stmt' {
                $stats.Stmt++
                $stack = $st.Stack
                if ($stack.Count -eq 0) {
                    [void](Push-AlFrame -State $st -Ev $ev -StartTicks $ev.Ticks -Synthetic $true -Stats $stats)
                    $stats.OrphanFrames++
                }
                else {
                    $top = $stack[$stack.Count - 1]
                    if ($top.FunctionName -ne $ev.FunctionName -or $top.ObjectId -ne $ev.ObjectId -or
                        $top.ObjectType   -ne $ev.ObjectType) {
                        # рамка своей функции ниже по стеку — значит потерян выход
                        $found = -1
                        for ($k = $stack.Count - 2; $k -ge 0; $k--) {
                            if ($stack[$k].FunctionName -eq $ev.FunctionName -and
                                $stack[$k].ObjectId     -eq $ev.ObjectId -and
                                $stack[$k].ObjectType   -eq $ev.ObjectType) { $found = $k; break }
                        }
                        if ($found -ge 0) {
                            while (($stack.Count - 1) -gt $found) {
                                Pop-AlFrame -State $st -EndTicks $ev.Ticks -Stats $stats
                                $stats.ForcedPops++
                            }
                        }
                        else {
                            [void](Push-AlFrame -State $st -Ev $ev -StartTicks $ev.Ticks -Synthetic $true -Stats $stats)
                            $stats.MissedStarts++
                        }
                    }
                }

                $top = $stack[$stack.Count - 1]
                Close-AlFrameStatement -Frame $top -EndTicks $ev.Ticks -State $st -Stats $stats

                $line = -1
                if ($LineSource -ne 'Offset' -and $FallbackMap) {
                    $norm = (ConvertTo-AlNormText $ev.Statement).ToUpperInvariant()
                    $key  = $ev.FunctionName + "`t" + $norm
                    if ($fbMemo.ContainsKey($key)) { $line = $fbMemo[$key] }
                    else {
                        if ($FallbackMap.ContainsKey($key)) { $line = [int]$FallbackMap[$key] }
                        elseif ($fbIndex.ContainsKey($ev.FunctionName)) {
                            $bestLen = 0
                            foreach ($c in $fbIndex[$ev.FunctionName]) {
                                if ($c.Text.Length -le $bestLen -or $c.Text.Length -ge $norm.Length) { continue }
                                if ($norm.StartsWith($c.Text, 'Ordinal')) { $bestLen = $c.Text.Length; $line = $c.LineNo }
                            }
                        }
                        $fbMemo[$key] = $line
                    }
                }
                if ($line -lt 0 -and $LineSource -ne 'Fallback') { $line = $ev.LineNumber + $LineOffset }

                if ($line -lt 0) { $top.StmtBucket = $null; $stats.Unresolved++ }
                else {
                    $top.StmtBucket = Get-AlLineBucket -Buckets $buckets -ObjectType $ev.ObjectType `
                                                       -ObjectId $ev.ObjectId -LineNo $line
                }
                $top.HasStmt     = $true
                $top.StmtStart   = $ev.Ticks
                $top.StmtChild   = 0L
                $top.StmtCallees = 0
                $top.StmtText    = $ev.Statement
            }

            'Start' {
                $stats.Start++
                [void](Push-AlFrame -State $st -Ev $ev -StartTicks $ev.Ticks -Synthetic $false -Stats $stats)
            }

            'Stop' {
                $stats.Stop++
                $stack = $st.Stack
                if ($stack.Count -eq 0) { $stats.UnmatchedStops++; continue }
                $top = $stack[$stack.Count - 1]
                if ($top.FunctionName -ne $ev.FunctionName -or $top.ObjectId -ne $ev.ObjectId -or
                    $top.ObjectType   -ne $ev.ObjectType) {
                    $found = -1
                    for ($k = $stack.Count - 2; $k -ge 0; $k--) {
                        if ($stack[$k].FunctionName -eq $ev.FunctionName -and
                            $stack[$k].ObjectId     -eq $ev.ObjectId -and
                            $stack[$k].ObjectType   -eq $ev.ObjectType) { $found = $k; break }
                    }
                    if ($found -lt 0) { $stats.UnmatchedStops++; continue }
                    while (($stack.Count - 1) -gt $found) {
                        Pop-AlFrame -State $st -EndTicks $ev.Ticks -Stats $stats
                        $stats.ForcedPops++
                    }
                }
                Pop-AlFrame -State $st -EndTicks $ev.Ticks -Stats $stats
            }

            'Error' {
                $stats.Error++
                $stack = $st.Stack
                if ($stack.Count -eq 0) { continue }
                # Функция, отработавшая с ошибкой, ЗАВЕРШИЛАСЬ: парного ALFunctionStop у
                # неё не будет, платформа шлёт вместо него ALFunctionError - «failed»
                # вместо «exited normally». Раньше рамка оставалась на стеке до конца
                # разбора, и в Total строки вызова втягивалось всё, что случилось в
                # сессии после ошибки: хвост закрывал её последней меткой сессии.
                # Поэтому ошибка сворачивает стек так же, как выход.
                #
                # Перехваченная ошибка этим не ломается: вызывающий продолжает работу
                # своими событиями оператора, его рамка ниже по стеку и не трогается.
                $found = -1
                for ($k = $stack.Count - 1; $k -ge 0; $k--) {
                    if ($stack[$k].FunctionName -eq $ev.FunctionName -and
                        $stack[$k].ObjectId     -eq $ev.ObjectId -and
                        $stack[$k].ObjectType   -eq $ev.ObjectType) { $found = $k; break }
                }
                if ($found -lt 0) {
                    # Своей рамки на стеке нет - закрываем хотя бы открытый оператор
                    # верхней, иначе его Total дотянется до следующего события.
                    $stats.ErrorUnmatched++
                    Close-AlFrameStatement -Frame $stack[$stack.Count - 1] -EndTicks $ev.Ticks -State $st -Stats $stats
                    continue
                }
                while ($stack.Count -gt $found) {
                    Pop-AlFrame -State $st -EndTicks $ev.Ticks -Stats $stats
                    $stats.ErrorPops++
                }
            }

            'PairStart' {
                $stats.SqlStart++
                $stack = $st.Stack
                $b = $null; $fr = $null
                if ($stack.Count -gt 0) { $fr = $stack[$stack.Count - 1]; $b = $fr.StmtBucket }
                if ($null -eq $b) { $stats.SqlOrphan++ }
                if ($st.Sql.Count -gt 0) { $stats.SqlNested++ }
                # Рамка запоминается, чтобы операцию было чем снять, если Stop потерялся:
                # закрытие оператора этой рамки её и уберёт (Close-AlFrameStatement).
                [void]$st.Sql.Add(@{ Op = $ev.OpName; Ticks = $ev.Ticks; Bucket = $b; Frame = $fr })
            }

            'PairStop' {
                $stats.SqlStop++
                $q  = $st.Sql
                $ix = -1
                for ($k = $q.Count - 1; $k -ge 0; $k--) { if ($q[$k].Op -eq $ev.OpName) { $ix = $k; break } }
                if ($ix -lt 0) { $stats.SqlUnmatched++; continue }
                $o = $q[$ix]
                $q.RemoveAt($ix)
                $d = $ev.Ticks - $o.Ticks
                if ($d -lt 0) { $d = 0L }
                # Вложенность определяется ПОЛОЖЕНИЕМ в списке открытых операций, а не
                # отметкой, снятой в момент начала: после снятия огрызков список сам себя
                # чинит, а запомненная когда-то глубина осталась бы завышенной навсегда.
                # Нулевая позиция - операция самая внешняя из открытых, её время и идёт
                # в зачёт; вложенная не удваивает время, наружная уже включает её.
                $outer = ($ix -eq 0)
                if ($null -ne $o.Bucket) {
                    $o.Bucket.SqlCount++
                    if ($outer) { $o.Bucket.SqlTicks += $d }
                }
                $stats.SqlCount++
                if ($outer) { $stats.SqlTicks += $d }
            }

            default { $stats.Other++ }
        }
    }

    # хвост: всё, что осталось открытым, закрывается последней меткой своей сессии
    foreach ($sid in @($state.Keys)) {
        $st = $state[$sid]
        while ($st.Stack.Count -gt 0) {
            Pop-AlFrame -State $st -EndTicks $st.Last -Stats $stats
            $stats.UnclosedFrames++
        }
        # Операции, начатые вообще без рамки: снятие по рамке до них не добирается.
        if ($st.Sql.Count -gt 0) { $stats.SqlDropped += $st.Sql.Count; $st.Sql.Clear() }
        $stats.RootTicks += $st.RootTicks
        $stats.SelfTicks += $st.SelfTicks
        $stats.Sessions[$sid] = [pscustomobject]@{
            SessionId  = $sid;         Events    = $st.Events
            FirstTicks = $st.First;    LastTicks = $st.Last
            RootTicks  = $st.RootTicks; SelfTicks = $st.SelfTicks
        }
    }

    # текст строки из листинга
    $text = $null
    if ($Listing) {
        $text = @{}
        foreach ($l in $Listing) { $text[$l.LineNo] = $l.Text }
    }

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($k in $buckets.Keys) {
        $b = $buckets[$k]
        if ($ObjectType -ge 0 -and $b.ObjectType -ne $ObjectType) { continue }
        if ($ObjectId   -ge 0 -and $b.ObjectId   -ne $ObjectId)   { continue }
        $min = $b.MinTicks
        if ($min -eq [int64]::MaxValue) { $min = 0L }
        $avg = 0.0
        if ($b.Hits -gt 0) { $avg = $b.TotalTicks / [double]$b.Hits }
        $t = ''
        if ($text -and $text.ContainsKey($b.LineNo)) { $t = $text[$b.LineNo] }
        [void]$out.Add([pscustomobject]@{
            ObjectType    = $b.ObjectType
            ObjectId      = $b.ObjectId
            LineNo        = $b.LineNo
            FunctionName  = $b.FunctionName
            Hits          = $b.Hits
            TotalTicks    = $b.TotalTicks
            SelfTicks     = $b.SelfTicks
            ChildTicks    = $b.ChildTicks
            SqlTicks      = $b.SqlTicks
            SqlCount      = $b.SqlCount
            MinTicks      = $min
            MaxTicks      = $b.MaxTicks
            TotalMs       = [math]::Round($b.TotalTicks / $script:AlTicksPerMs, 3)
            SelfMs        = [math]::Round($b.SelfTicks  / $script:AlTicksPerMs, 3)
            ChildMs       = [math]::Round($b.ChildTicks / $script:AlTicksPerMs, 3)
            SqlMs         = [math]::Round($b.SqlTicks   / $script:AlTicksPerMs, 3)
            SelfNoSqlMs   = [math]::Round(($b.SelfTicks - $b.SqlTicks) / $script:AlTicksPerMs, 3)
            AvgMs         = [math]::Round($avg / $script:AlTicksPerMs, 4)
            MinMs         = [math]::Round($min / $script:AlTicksPerMs, 3)
            MaxMs         = [math]::Round($b.MaxTicks / $script:AlTicksPerMs, 3)
            Callees       = $b.Callees
            Stmts         = $b.Texts.Count
            Recursive     = $b.Recursive
            RecursiveHits = $b.RecursiveHits
            Text          = $t
        })
    }

    $stats.Lines           = $out.Count
    $stats.AllLines        = $buckets.Count
    $script:AlMeasureStats = $stats

    return ($out | Sort-Object ObjectType, ObjectId, LineNo)
}

function Export-AlLineMetrics {
    <#
    .SYNOPSIS
        Пишет результат Measure-AlLines в TSV метрик, который читает Build-Report.ps1.

    .DESCRIPTION
        Отдельная функция нужна из-за формата чисел. Build-Report разбирает TSV
        построчно, по разделителю табуляции, и переводит значения в число по
        инвариантной культуре. Значит Export-Csv не годится вовсе (в PowerShell 5.1
        он берёт каждое поле в кавычки), а обычное форматирование в русской локали
        подставило бы разделитель разрядов. Здесь числа пишутся инвариантно и без
        разрядов, файл — UTF-8 без BOM с CRLF, как и events.tsv.

        Колонки, которые читает отчёт: ObjectType, ObjectId, LineNo, Hits, TotalMs,
        SelfMs, SqlCount, SqlMs. Остальные пишутся про запас и отчёту не мешают.

    .PARAMETER Lines
        Результат Measure-AlLines.

    .PARAMETER Path
        Куда писать (обычно out\lines.tsv).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]] $Lines,
        [Parameter(Mandatory)][string]   $Path
    )

    $inv  = [System.Globalization.CultureInfo]::InvariantCulture
    $cols = @('ObjectType','ObjectId','LineNo','FunctionName','Hits','TotalMs','SelfMs','ChildMs',
              'SqlMs','SelfNoSqlMs','AvgMs','MinMs','MaxMs','SqlCount','Callees','Stmts',
              'Recursive','RecursiveHits','TotalTicks','SelfTicks','SqlTicks')
    $tab  = [char]9

    $rows = New-Object System.Collections.Generic.List[string]
    [void]$rows.Add(($cols -join $tab))
    foreach ($r in $Lines) {
        $v = New-Object System.Collections.Generic.List[string]
        foreach ($c in $cols) {
            $x = $r.$c
            if ($null -eq $x) { [void]$v.Add(''); continue }
            if ($x -is [double] -or $x -is [decimal] -or $x -is [single]) {
                [void]$v.Add(([double]$x).ToString('0.####', $inv))
            }
            elseif ($x -is [bool]) { [void]$v.Add($(if ($x) { '1' } else { '0' })) }
            else { [void]$v.Add(([string]$x).Replace($tab, ' ')) }
        }
        [void]$rows.Add(($v -join $tab))
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, (($rows -join "`r`n") + "`r`n"), $utf8NoBom)
    return $Path
}

# ---------------------------------------------------------------------------
# 4. инварианты прогона
# ---------------------------------------------------------------------------

function Get-AlRunSummary {
    <#
    .SYNOPSIS
        Сводка валидности прогона: wall-clock, измеренное время, «учтено времени %»,
        события, потери, сессии, диапазон времени.

    .DESCRIPTION
        Три разных времени, которые нельзя путать:
          WallMs     — от первого события до последнего, вместе с простоем между
                       действиями человека;
          MeasuredMs — сумма длительностей КОРНЕВЫХ рамок вызова, то есть время,
                       когда сервер действительно исполнял C/AL;
          SelfMs     — сумма Self всех строк.
        «Учтено времени» = SelfMs / MeasuredMs: какая доля измеренного времени легла
        на конкретные строки. Именно этот инвариант ловит потерю событий и ошибки
        разбора, и он же гейтит прогон (ниже -MinAccountedPct — невалиден).
        Отношение MeasuredMs / WallMs даётся отдельно (CoveragePct) и НЕ гейтит: на
        живом сценарии человек думает между кликами, и низкое покрытие wall-clock там
        нормально.

        Потери событий видны по разрыву нумерации записей канала (RecordId). Второй
        вид потерь — переполнение канала — в RecordId не виден вовсе и ловится
        сборщиком по размеру файла журнала.

        Нумерация есть НЕ ВСЕГДА. Это свойство канала, а не наше: у записей `.etl`
        (режим Full — основной) EventRecord.RecordId пуст, номера записи там просто
        нет, и колонка в events.tsv выходит пустой. У `.evtx` (режим Lite читает
        Application) номер есть. Раньше сводка в обоих случаях печатала «Потеряно: 0»
        — утверждение, которого данные не подтверждают. Теперь при отсутствии
        нумерации проверка объявляется НЕ ПРОВЕДЁННОЙ; для `.etl` потери ловит
        сборщик по счётчику потерянных буферов ядра. Частичная нумерация тоже не
        считается: у неё «разрыв» получался бы на каждом ненумерованном событии.

        ВАЖНО: -Events подавать ЦЕЛИКОМ, без фильтра по сессии. Нумерация записей
        сквозная по каналу, поэтому на отфильтрованном наборе разрывы будут ложными
        и прогон объявится невалидным на ровном месте. Метрики времени при этом можно
        брать от отфильтрованного разбора: они приходят из -MeasureStats.

    .PARAMETER Events
        Полный набор событий из Import-AlTraceEvents, без фильтра по сессии.

    .PARAMETER MeasureStats
        Диагностика от Measure-AlLines; по умолчанию берётся от последнего вызова.

    .PARAMETER MinAccountedPct
        Порог «учтено времени», ниже которого прогон объявляется невалидным.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]] $Events,
        [hashtable] $MeasureStats,
        [double]    $MinAccountedPct = 90.0
    )

    if (-not $MeasureStats) { $MeasureStats = $script:AlMeasureStats }

    $n = 0
    $tMin = [int64]::MaxValue; $tMax = [int64]::MinValue
    $rMin = [int64]::MaxValue; $rMax = [int64]::MinValue; $rCnt = 0
    $sessions = @{}
    $byKind   = @{}
    foreach ($e in $Events) {
        $n++
        if ($e.Ticks -gt 0) {
            if ($e.Ticks -lt $tMin) { $tMin = $e.Ticks }
            if ($e.Ticks -gt $tMax) { $tMax = $e.Ticks }
        }
        if ($e.RecordId -gt 0) {
            $rCnt++
            if ($e.RecordId -lt $rMin) { $rMin = $e.RecordId }
            if ($e.RecordId -gt $rMax) { $rMax = $e.RecordId }
        }
        $sessions[$e.SessionId] = $true
        if ($byKind.ContainsKey($e.Kind)) { $byKind[$e.Kind]++ } else { $byKind[$e.Kind] = 1 }
    }
    if ($n -eq 0 -or $tMin -eq [int64]::MaxValue) { $tMin = 0L; $tMax = 0L }

    $wall = 0L
    if ($tMax -ge $tMin) { $wall = $tMax - $tMin }
    # Считать разрыв можно, только если пронумерованы ВСЕ события: на частично
    # пронумерованном наборе каждое событие без номера выглядело бы дырой.
    $lost        = 0L
    $lossChecked = ($n -gt 0 -and $rCnt -eq $n -and $rMax -ge $rMin)
    if ($lossChecked) {
        $lost = ($rMax - $rMin + 1) - $rCnt
        if ($lost -lt 0) { $lost = 0L }
    }

    $root = 0L; $self = 0L; $sqlT = 0L; $sqlN = 0L; $depth = 0; $lines = 0
    if ($MeasureStats) {
        $root  = [int64]$MeasureStats.RootTicks
        $self  = [int64]$MeasureStats.SelfTicks
        $sqlT  = [int64]$MeasureStats.SqlTicks
        $sqlN  = [int64]$MeasureStats.SqlCount
        $depth = [int]$MeasureStats.MaxDepth
        $lines = [int]$MeasureStats.AllLines
    }

    $accounted = 0.0
    if ($root -gt 0) { $accounted = 100.0 * $self / $root }
    $coverage = 0.0
    if ($wall -gt 0) { $coverage = 100.0 * $root / $wall }

    $blockers = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $hasStmt = $byKind.ContainsKey('Stmt')
    if ($n -eq 0) { [void]$blockers.Add('событий нет вовсе') }
    elseif (-not $hasStmt) {
        [void]$blockers.Add('нет событий уровня оператора: трассировка не была включена')
    }
    # разрыв нумерации гейтит только настоящий построчный трейс: в общем журнале
    # Application лежат записи чужих провайдеров, и «дыры» там штатны
    if ($lost -gt 0) {
        if ($hasStmt) { [void]$blockers.Add(('разрыв нумерации записей канала: {0:N0}' -f $lost)) }
        else { [void]$warnings.Add(('разрыв нумерации записей канала: {0:N0} (журнал общий — не показатель потерь)' -f $lost)) }
    }
    # Отсутствие нумерации - не потеря и не норма, а непроверенное место: молчать
    # о нём нельзя, объявлять прогон невалидным - тем более (у .etl так всегда).
    if (-not $lossChecked -and $n -gt 0) {
        if ($rCnt -eq 0) {
            [void]$warnings.Add('нумерации записей канала нет: потери по разрыву не проверялись ' +
                                '(в .etl её не бывает — переполнение канала ловит сборщик по счётчику ядра)')
        }
        else {
            # лишние скобки не для красоты: в аргументе метода запятая разделяет
            # АРГУМЕНТЫ, и без них $n уходил вторым параметром .Add()
            [void]$warnings.Add((('нумерация записей канала есть не у всех событий ({0:N0} из {1:N0}): ' +
                                  'потери по разрыву не проверялись') -f $rCnt, $n))
        }
    }
    if ($root -gt 0 -and $accounted -lt $MinAccountedPct) {
        [void]$blockers.Add(('учтено {0:N1} % измеренного времени при пороге {1:N0} %' -f $accounted, $MinAccountedPct))
    }
    if ($root -eq 0 -and $n -gt 0 -and $byKind.ContainsKey('Stmt')) {
        [void]$blockers.Add('не закрылась ни одна корневая рамка вызова')
    }

    if ($MeasureStats) {
        if ($MeasureStats.UnmatchedStops  -gt 0) { [void]$warnings.Add(('выходов без входа: {0}' -f $MeasureStats.UnmatchedStops)) }
        if ($MeasureStats.MissedStarts    -gt 0) { [void]$warnings.Add(('операторов без входа в функцию: {0}' -f $MeasureStats.MissedStarts)) }
        if ($MeasureStats.OrphanFrames    -gt 0) { [void]$warnings.Add(('рамок начато с середины: {0}' -f $MeasureStats.OrphanFrames)) }
        if ($MeasureStats.ForcedPops      -gt 0) { [void]$warnings.Add(('рамок свёрнуто принудительно: {0}' -f $MeasureStats.ForcedPops)) }
        if ($MeasureStats.UnclosedFrames  -gt 0) { [void]$warnings.Add(('рамок осталось открытыми: {0}' -f $MeasureStats.UnclosedFrames)) }
        if ($MeasureStats.RecursiveFrames -gt 0) { [void]$warnings.Add(('рекурсивных рамок: {0} — Total на их строках завышен' -f $MeasureStats.RecursiveFrames)) }
        if ($MeasureStats.NegativeSelf    -gt 0) { [void]$warnings.Add(('операторов с отрицательным Self: {0}' -f $MeasureStats.NegativeSelf)) }
        if ($MeasureStats.SqlUnmatched    -gt 0) { [void]$warnings.Add(('SQL-событий без пары: {0}' -f $MeasureStats.SqlUnmatched)) }
        if ($MeasureStats.SqlOrphan       -gt 0) { [void]$warnings.Add(('SQL вне оператора C/AL: {0}' -f $MeasureStats.SqlOrphan)) }
        if ($MeasureStats.SqlDropped      -gt 0) { [void]$warnings.Add(('SQL-операций без завершения: {0} — их время не учтено ни на одной строке' -f $MeasureStats.SqlDropped)) }
        if ($MeasureStats.ErrorPops       -gt 0) { [void]$warnings.Add(('рамок свёрнуто по ошибке C/AL: {0}' -f $MeasureStats.ErrorPops)) }
        if ($MeasureStats.ErrorUnmatched  -gt 0) { [void]$warnings.Add(('ошибок C/AL без своей рамки на стеке: {0}' -f $MeasureStats.ErrorUnmatched)) }
        if ($MeasureStats.Unresolved      -gt 0) { [void]$warnings.Add(('операторов без номера строки: {0}' -f $MeasureStats.Unresolved)) }
    }

    $first = $null; $last = $null
    if ($tMin -gt 0) { $first = [datetime]$tMin; $last = [datetime]$tMax }

    return [pscustomobject]@{
        Events       = $n
        ByKind       = $byKind
        Sessions     = $sessions.Count
        FirstTime    = $first
        LastTime     = $last
        WallMs       = [math]::Round($wall / $script:AlTicksPerMs, 1)
        MeasuredMs   = [math]::Round($root / $script:AlTicksPerMs, 1)
        SelfMs       = [math]::Round($self / $script:AlTicksPerMs, 1)
        SqlMs        = [math]::Round($sqlT / $script:AlTicksPerMs, 1)
        SqlCount     = $sqlN
        AccountedPct = [math]::Round($accounted, 2)
        CoveragePct  = [math]::Round($coverage, 2)
        LostEvents   = $lost
        LossChecked  = $lossChecked
        Numbered     = $rCnt
        MaxDepth     = $depth
        Lines        = $lines
        IsValid      = ($blockers.Count -eq 0)
        Blockers     = $blockers.ToArray()
        Warnings     = $warnings.ToArray()
        Stats        = $MeasureStats
    }
}

function Write-AlRunSummary {
    <#
    .SYNOPSIS
        Печать сводки прогона в консоль, компактно.
    #>
    param([Parameter(Mandatory)] $Summary)

    $k = $Summary.ByKind
    $get = { param($name) if ($k.ContainsKey($name)) { $k[$name] } else { 0 } }

    Write-Host ''
    Write-Host ('Событий:   {0:N0} в {1} сессиях; строк {2:N0}; глубина стека {3}' -f
        $Summary.Events, $Summary.Sessions, $Summary.Lines, $Summary.MaxDepth)
    Write-Host ('  оператор {0:N0}; вход {1:N0}; выход {2:N0}; SQL {3:N0}/{4:N0}; прочее {5:N0}' -f
        (& $get 'Stmt'), (& $get 'Start'), (& $get 'Stop'),
        (& $get 'PairStart'), (& $get 'PairStop'), (& $get 'Other'))
    if ($Summary.FirstTime) {
        Write-Host ('Период:    {0:HH:mm:ss.fff} .. {1:HH:mm:ss.fff}' -f $Summary.FirstTime, $Summary.LastTime)
    }
    Write-Host ('Время:     wall {0:N1} мс; измерено {1:N1} мс; по строкам {2:N1} мс; SQL {3:N1} мс' -f
        $Summary.WallMs, $Summary.MeasuredMs, $Summary.SelfMs, $Summary.SqlMs)
    Write-Host ('Учтено:    {0:N2} % измеренного времени (покрытие wall-clock {1:N2} %)' -f
        $Summary.AccountedPct, $Summary.CoveragePct)
    if ($Summary.LossChecked) { Write-Host ('Потеряно:  {0:N0}' -f $Summary.LostEvents) }
    else {
        Write-Host 'Потеряно:  не проверялось — у записей нет нумерации канала' -ForegroundColor DarkYellow
    }
    foreach ($w in $Summary.Warnings) { Write-Host ('  ! ' + $w) -ForegroundColor DarkYellow }
    if ($Summary.IsValid) { Write-Host 'Прогон:    валиден' -ForegroundColor Green }
    else {
        Write-Host 'Прогон:    НЕВАЛИДЕН' -ForegroundColor Red
        foreach ($b in $Summary.Blockers) { Write-Host ('  - ' + $b) -ForegroundColor Red }
    }
}
