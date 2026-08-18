#Requires -Version 5.1
<#
.SYNOPSIS
    Ядро статического разбора C/AL поверх дампа .alsrc. Подключается через dot-sourcing.

.DESCRIPTION
    Четыре функции — фундамент для движка подсказок (правила пишутся отдельно):

      Get-AlLexed        — «очищенный» текст каждой строки: строковые литералы заменены
                           на пустые '', комментарии (// и { }) вырезаны, пробелы схлопнуты.
                           Регулярки правил не срабатывают на текст внутри строк/комментариев.
      Get-AlSymbols      — таблица символов объекта: переменная -> тип; для Record —
                           номер таблицы и признак TEMPORARY; область (глобальная/функция).
      Get-AlStructure    — структура кода: границы функций, циклы REPEAT/FOR/WHILE, блоки
                           WITH; на каждую строку — глубина вложенности циклов и объемлющий
                           цикл. Если разбор функции не сошёлся — флаг Reliable=false.
      Get-AlFilterChains — цепочки «SETRANGE/SETFILTER/SETCURRENTKEY/RESET -> потребитель
                           (FINDSET/COUNT/...)» по Record-переменным в пределах функции.

    Источник объявлений переменных — текстовый экспорт baseline (сопоставление по типу и
    номеру объекта, номера строк не используются, поэтому расхождение нумерации экспорта
    и дампа не мешает). Альтернатива — BLOB [Object Metadata].[Metadata] — отвергнута:
    на рабочей базе он пуст у 1834 объектов с исходником, включая Codeunit 80 Sales-Post.
    Плата за выбор: если объект в базе перекомпилирован после выгрузки baseline,
    состав переменных может отставать; несовпадение областей с картой функций дампа —
    сигнал устаревания (проверяется потребителем).

    Требует Lib-AlListing.ps1 (Get-AlListing, Get-AlFunctionMap); подключает её сама.

    Подключение:  . (Join-Path $PSScriptRoot 'Lib-AlParse.ps1')
#>

# Подключается безусловно: повторный dot-source стоит миллисекунды, а проверка
# «функция уже есть» обманчива — функция может быть унаследована от вызывающего
# скрипта, тогда как переменные уровня файла в нашей области так и не появятся.
. (Join-Path $PSScriptRoot 'Lib-AlListing.ps1')

# --- вспомогательное --------------------------------------------------------

function ConvertTo-AlRows {
    <#
    .SYNOPSIS
        Приводит вход к плоскому массиву строк-объектов.

    .DESCRIPTION
        Функции Lib-AlListing возвращают коллекцию одним объектом (return ,$list),
        а вызывающий мог обернуть её в @(). Разворачивает оба варианта, чтобы
        индексация Rows[LineNo-1] всегда работала.
    #>
    param($Value)
    # ВАЖНО: НЕ @($Value) — в WinPS 5.1 оператор @() на List[object] из другой функции
    # падает с ArgumentException «Argument types do not match»; перебираем явно.
    $list = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Value) {
        if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
            foreach ($v in $Value) {
                if ($null -eq $v) { continue }
                if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string] -and
                    $v -isnot [System.Management.Automation.PSCustomObject]) {
                    foreach ($w in $v) { $list.Add($w) }   # вложенная коллекция: @() поверх ,$list
                }
                else { $list.Add($v) }
            }
        }
        else { $list.Add($Value) }
    }
    return ,$list.ToArray()
}

# --- справочники ------------------------------------------------------------

# Подпапка и префикс имени файла в каталоге baseline по типу объекта.
$script:AlBaseTypeDirs = @{
    1 = @('Tables',    't'); 3 = @('Reports',  'r'); 5 = @('Codeunits', 'c')
    6 = @('XMLports',  'x'); 7 = @('MenuSuites','m'); 8 = @('Pages',    'p')
    9 = @('Queries',   'q')
}

# Кэш «номер таблицы -> имя» из .alsrc\index.tsv (строки Type=1).
$script:AlTableNameCache = $null

function Get-AlBaseRoot {
    <#
    .SYNOPSIS
        Каталог baseline: -Root, иначе $env:LP_BASELINE_DIR, иначе <корень репозитория>\baseline.

    .DESCRIPTION
        Baseline — текстовый экспорт объектов целевой базы из C/SIDE, разложенный по
        подпапкам Tables, Pages, Codeunits и т.д. В репозиторий он не входит: у каждой
        установки он свой. Если каталога нет, объявления переменных недоступны — правила
        обязаны считать типы неизвестными, а не угадывать.
    #>
    param([string]$Root)
    if ($Root) { return $Root }
    if ($env:LP_BASELINE_DIR) { return $env:LP_BASELINE_DIR }
    return (Join-Path (Split-Path -Parent $PSScriptRoot) 'baseline')
}

function Get-AlTableNameById {
    <#
    .SYNOPSIS
        Имя таблицы по номеру: .alsrc\index.tsv, затем out\keys.tsv; '' если неизвестно.

    .DESCRIPTION
        Дамп .alsrc содержит только скомпилированные объекты; таблицу, известную лишь
        по экспорту baseline (например, Sales Line 37), добирает индекс ключей
        Build-KeysIndex.ps1 (<корень репозитория>\out\keys.tsv).
    #>
    param([int]$TableId, [string]$SourceRoot)
    if ($null -eq $script:AlTableNameCache) {
        $script:AlTableNameCache = @{}
        $idx = Join-Path (Get-AlSourceRoot $SourceRoot) 'index.tsv'
        if (Test-Path -LiteralPath $idx) {
            $lines = [System.IO.File]::ReadAllLines($idx, [System.Text.Encoding]::UTF8)
            for ($i = 1; $i -lt $lines.Length; $i++) {
                $c = $lines[$i] -split "`t"
                if ($c.Length -ge 4 -and $c[0] -eq '1') { $script:AlTableNameCache[[int]$c[2]] = $c[3] }
            }
        }
        $keys = Join-Path (Split-Path -Parent $PSScriptRoot) 'out\keys.tsv'
        if (Test-Path -LiteralPath $keys) {
            $lines = [System.IO.File]::ReadAllLines($keys, [System.Text.Encoding]::UTF8)
            for ($i = 1; $i -lt $lines.Length; $i++) {
                $c = $lines[$i] -split "`t"
                if ($c.Length -ge 2) {
                    $id = [int]$c[0]
                    if (-not $script:AlTableNameCache.ContainsKey($id)) { $script:AlTableNameCache[$id] = $c[1] }
                }
            }
        }
    }
    if ($script:AlTableNameCache.ContainsKey($TableId)) { return $script:AlTableNameCache[$TableId] }
    return ''
}

function Find-AlBaseExport {
    <#
    .SYNOPSIS
        Путь к файлу текстового экспорта объекта в каталоге baseline (или $null).
    #>
    param(
        [Parameter(Mandatory)][int]$ObjectType,
        [Parameter(Mandatory)][int]$ObjectId,
        [string]$BaseRoot
    )
    if (-not $script:AlBaseTypeDirs.ContainsKey($ObjectType)) { return $null }
    $dir    = Join-Path (Get-AlBaseRoot $BaseRoot) $script:AlBaseTypeDirs[$ObjectType][0]
    $prefix = $script:AlBaseTypeDirs[$ObjectType][1]
    if (-not (Test-Path -LiteralPath $dir)) { return $null }
    $hit = @(Get-ChildItem -LiteralPath $dir -Filter ("{0}{1} - *.txt" -f $prefix, $ObjectId) |
             Select-Object -First 1)
    if ($hit.Count -gt 0) { return $hit[0].FullName }
    return $null
}

# === Get-AlLexed ============================================================

function Get-AlLexed {
    <#
    .SYNOPSIS
        «Очищенный» текст каждой строки листинга: без строк и комментариев.

    .DESCRIPTION
        На каждую строку листинга (позиционно, Lexed[i] соответствует Listing[i]):
            LineNo, Kind, FunctionName — как в листинге;
            Clean — текст без строковых литералов (заменены на ''), без //-комментариев
                    и без { }-комментариев (бывают многострочными), пробелы схлопнуты.
        Правила C/AL: строка — в одинарных кавычках, '' внутри — сама кавычка; имена
        в двойных кавычках — идентификаторы, их содержимое сохраняется.

        Возвращает коллекцию одним объектом (как Lib-AlListing): не оборачивайте в @().
    #>
    param([Parameter(Mandatory)][object]$Listing)
    $rows = ConvertTo-AlRows $Listing

    $result  = New-Object System.Collections.Generic.List[object]
    $inBrace = $false
    $wsRx    = [regex]'\s+'

    foreach ($row in $rows) {
        $text  = $row.Text
        $clean = ''
        if ($row.Kind -eq 'Empty' -or $row.Kind -eq 'Comment') {
            if ($inBrace -and $text.IndexOf('}') -ge 0) {
                # закрытие { }-комментария может прятаться и в «комментарной» строке
                $inBrace = $false
                $rest = $text.Substring($text.IndexOf('}') + 1)
                if ($rest.TrimStart().StartsWith('//')) { $rest = '' }
                $clean = $wsRx.Replace($rest, ' ').Trim()
            }
        }
        elseif (-not $inBrace -and $text.IndexOf("'") -lt 0 -and $text.IndexOf('{') -lt 0 -and
                $text.IndexOf('}') -lt 0 -and $text.IndexOf('//') -lt 0) {
            # быстрый путь: спецсимволов нет
            $clean = $wsRx.Replace($text, ' ').Trim()
        }
        else {
            $sb = New-Object System.Text.StringBuilder
            $i = 0; $n = $text.Length
            while ($i -lt $n) {
                $ch = $text[$i]
                if ($inBrace) {
                    if ($ch -eq '}') { $inBrace = $false }
                    $i++
                    continue
                }
                if ($ch -eq "'") {
                    # строковый литерал: заменяем на пустой '', '' внутри — экранирование
                    [void]$sb.Append("''")
                    $i++
                    while ($i -lt $n) {
                        if ($text[$i] -eq "'") {
                            if ($i + 1 -lt $n -and $text[$i + 1] -eq "'") { $i += 2; continue }
                            $i++
                            break
                        }
                        $i++
                    }
                    continue
                }
                if ($ch -eq '"') {
                    # имя в двойных кавычках — идентификатор, сохраняем как есть
                    [void]$sb.Append('"')
                    $i++
                    while ($i -lt $n) {
                        [void]$sb.Append($text[$i])
                        if ($text[$i] -eq '"') { $i++; break }
                        $i++
                    }
                    continue
                }
                if ($ch -eq '/' -and $i + 1 -lt $n -and $text[$i + 1] -eq '/') { break }
                if ($ch -eq '{') { $inBrace = $true; $i++; continue }
                [void]$sb.Append($ch)
                $i++
            }
            $clean = $wsRx.Replace($sb.ToString(), ' ').Trim()
        }

        $result.Add([pscustomobject]@{
            LineNo       = $row.LineNo
            Kind         = $row.Kind
            FunctionName = $row.FunctionName
            Clean        = $clean
        })
    }
    return ,$result
}

# === Get-AlSymbols ==========================================================

function ConvertTo-AlSymbol {
    <#
    .SYNOPSIS
        Разбор одного объявления 'Имя@1000 : [VAR] [TEMPORARY] Тип' в символ.
    #>
    param(
        [Parameter(Mandatory)][string]$Decl,
        [string]$Scope  = '',
        [string]$Origin = 'Local'
    )
    $s = $Decl.Trim()
    $byRef = $false
    if ($s -cmatch '^VAR\s+') { $byRef = $true; $s = $s.Substring(4).Trim() }
    $m = [regex]::Match($s, '^(?<name>"[^"]+"|[^@\s"]+)@[0-9]+\s*:\s*(?<type>.+)$')
    if (-not $m.Success) { return $null }

    $name = $m.Groups['name'].Value.Trim('"')
    $type = $m.Groups['type'].Value.Trim().TrimEnd(';').Trim()

    $temp = $false; $isArr = $false
    if ($type -cmatch '^TEMPORARY\s+') { $temp = $true; $type = $type.Substring(10).Trim() }
    $ma = [regex]::Match($type, '^ARRAY\s*\[[^\]]*\]\s*OF\s+')
    if ($ma.Success) { $isArr = $true; $type = $type.Substring($ma.Length).Trim() }
    if ($type -cmatch '^TEMPORARY\s+') { $temp = $true; $type = $type.Substring(10).Trim() }

    $kind = 'Simple'; $objId = $null
    $mk = [regex]::Match($type, '^(?<k>Record|Codeunit|Page|Form|Report|XMLport|Query|TestPage)\s+(?<id>[0-9]+)')
    if ($mk.Success) {
        $kind  = $mk.Groups['k'].Value
        $objId = [int]$mk.Groups['id'].Value
    }

    return [pscustomobject]@{
        Name      = $name
        Scope     = $Scope        # '' — глобальная, иначе имя функции/триггера как в дампе
        Origin    = $Origin       # Global | Local | Param
        Kind      = $kind         # Record | Codeunit | Page | ... | Simple
        ObjectId  = $objId        # для Record — номер таблицы
        TableName = ''
        Temporary = $temp
        ByRef     = $byRef
        IsArray   = $isArr
        DataType  = $type
    }
}

function Split-AlParams {
    <#
    .SYNOPSIS
        Делит список параметров по ';' верхнего уровня (кавычки и скобки учитываются).
    #>
    param([string]$Text)
    $parts = New-Object System.Collections.Generic.List[string]
    $depth = 0; $inSq = $false; $inDq = $false
    $start = 0
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($inSq) { if ($ch -eq "'") { $inSq = $false }; continue }
        if ($inDq) { if ($ch -eq '"') { $inDq = $false }; continue }
        switch ($ch) {
            "'" { $inSq = $true }
            '"' { $inDq = $true }
            '[' { $depth++ }
            ']' { if ($depth -gt 0) { $depth-- } }
            '(' { $depth++ }
            ')' { if ($depth -gt 0) { $depth-- } }
            ';' {
                if ($depth -eq 0) {
                    $parts.Add($Text.Substring($start, $i - $start))
                    $start = $i + 1
                }
            }
        }
    }
    if ($start -lt $Text.Length) { $parts.Add($Text.Substring($start)) }
    return $parts
}

function Get-AlParenSpan {
    <#
    .SYNOPSIS
        Содержимое первой (...) с учётом кавычек и вложенности; $null если не закрыта.
    #>
    param([string]$Text)
    $open = -1; $depth = 0; $inSq = $false; $inDq = $false
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($inSq) { if ($ch -eq "'") { $inSq = $false }; continue }
        if ($inDq) { if ($ch -eq '"') { $inDq = $false }; continue }
        if ($ch -eq "'") { $inSq = $true; continue }
        if ($ch -eq '"') { $inDq = $true; continue }
        if ($ch -eq '(') { if ($depth -eq 0) { $open = $i }; $depth++; continue }
        if ($ch -eq ')') {
            $depth--
            if ($depth -eq 0 -and $open -ge 0) {
                return $Text.Substring($open + 1, $i - $open - 1)
            }
        }
    }
    return $null
}

function Get-AlSymbols {
    <#
    .SYNOPSIS
        Таблица символов объекта: переменная -> тип, из текстового экспорта baseline.

    .DESCRIPTION
        Источники объявлений в экспорте:
          - глобальный VAR секции CODE;
          - параметры и локальный VAR каждой PROCEDURE/EVENT;
          - триггеры объекта в PROPERTIES (OnRun, OnInsert, ...) — область = имя триггера;
          - триггеры полей таблицы в FIELDS (OnValidate/OnLookup) — область
            '<Имя поля> - OnValidate', как называет функцию дамп .alsrc.
        НЕ разбираются секции CONTROLS/ACTIONS/DATASET (триггеры контролов страниц и
        датаитемов отчётов): для них вернутся только глобальные переменные.

        Если экспорта baseline нет (объект новее выгрузки) — вернётся пустой
        список с предупреждением: правила должны считать типы неизвестными.
    #>
    param(
        [Parameter(Mandatory)][int]$ObjectType,
        [Parameter(Mandatory)][int]$ObjectId,
        [string]$BaseRoot,
        [string]$SourceRoot
    )
    $out  = New-Object System.Collections.Generic.List[object]
    $path = Find-AlBaseExport -ObjectType $ObjectType -ObjectId $ObjectId -BaseRoot $BaseRoot
    if (-not $path) {
        Write-Warning ("Get-AlSymbols: нет экспорта baseline (LP_BASELINE_DIR) для {0} {1} — символы недоступны" -f `
            (Get-AlTypeName $ObjectType), $ObjectId)
        return ,$out
    }
    $lines = [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)

    # -- границы верхнеуровневых секций (PROPERTIES, FIELDS, CODE, ...) ------
    $sections = @{}
    $order = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i] -cmatch '^  ([A-Z][A-Z-]+)\s*$') {
            $order.Add(@{ Name = $Matches[1]; Start = $i })
        }
    }
    for ($i = 0; $i -lt $order.Count; $i++) {
        $end = $lines.Length - 1
        if ($i -lt $order.Count - 1) { $end = $order[$i + 1].Start - 1 }
        $sections[$order[$i].Name] = @{ Start = $order[$i].Start; End = $end }
    }

    $declRx    = [regex]'^\s*(?:VAR\s+)?("[^"]+"|[^@\s"]+)@[0-9]+\s*:\s*.+$'
    $trigRx    = [regex]'(?<t>On[A-Za-z]+)=VAR\s*$'
    $fieldRx   = [regex]'^\s*\{\s*[0-9]+\s*;[^;]*;(?<f>[^;]+);'

    # -- триггеры в PROPERTIES (объект) и FIELDS (поля таблицы) --------------
    foreach ($segName in @('PROPERTIES', 'FIELDS')) {
        if (-not $sections.ContainsKey($segName)) { continue }
        $seg = $sections[$segName]
        $field = ''
        for ($i = $seg.Start; $i -le $seg.End; $i++) {
            $ln = $lines[$i]
            if ($segName -eq 'FIELDS') {
                $mf = $fieldRx.Match($ln)
                if ($mf.Success) { $field = $mf.Groups['f'].Value.Trim() }
            }
            $mt = $trigRx.Match($ln)
            if (-not $mt.Success) { continue }
            $scope = $mt.Groups['t'].Value
            if ($segName -eq 'FIELDS' -and $field) { $scope = "$field - $scope" }
            $j = $i + 1
            while ($j -le $seg.End -and $lines[$j].Trim() -cne 'BEGIN') {
                if ($declRx.IsMatch($lines[$j])) {
                    $sym = ConvertTo-AlSymbol -Decl $lines[$j] -Scope $scope -Origin 'Local'
                    if ($sym) { $out.Add($sym) }
                }
                $j++
            }
            $i = $j
        }
    }

    # -- секция CODE: глобальный VAR, процедуры (параметры + локальный VAR) --
    if ($sections.ContainsKey('CODE')) {
        $seg = $sections['CODE']
        $headRx = [regex]'^    (?:LOCAL\s+)?(?:PROCEDURE|EVENT)\s'
        $seenProc = $false
        $i = $seg.Start
        while ($i -le $seg.End) {
            $ln = $lines[$i]

            if (-not $seenProc -and $ln -cmatch '^    VAR\s*$') {
                # глобальные переменные
                $j = $i + 1
                while ($j -le $seg.End -and $lines[$j] -match '^\s{5,}\S') {
                    if ($declRx.IsMatch($lines[$j])) {
                        $sym = ConvertTo-AlSymbol -Decl $lines[$j] -Scope '' -Origin 'Global'
                        if ($sym) { $out.Add($sym) }
                    }
                    $j++
                }
                $i = $j
                continue
            }

            if ($headRx.IsMatch($ln)) {
                $seenProc = $true
                # сигнатура может быть длинной, но в экспорте она на одной строке;
                # на всякий случай дочитываем до закрытой скобки и завершающей ';'
                $sig = $ln.Trim()
                $j = $i
                $guard = 0
                while ((($null -eq (Get-AlParenSpan $sig)) -or (-not $sig.EndsWith(';'))) -and
                       $j -lt $seg.End -and $guard -lt 10) {
                    $j++; $guard++
                    $sig = $sig + ' ' + $lines[$j].Trim()
                }
                $mn = [regex]::Match($sig, '^(?:LOCAL\s+)?(?:PROCEDURE|EVENT)\s+(?<name>"[^"]+"|[^@\s(]+)@')
                if ($mn.Success) {
                    $procName = $mn.Groups['name'].Value.Trim('"')
                    $paramTxt = Get-AlParenSpan $sig
                    if ($paramTxt) {
                        foreach ($p in (Split-AlParams $paramTxt)) {
                            if ($p.Trim().Length -eq 0) { continue }
                            $sym = ConvertTo-AlSymbol -Decl $p -Scope $procName -Origin 'Param'
                            if ($sym) { $out.Add($sym) }
                        }
                    }
                    # локальный VAR и тело
                    $j++
                    while ($j -le $seg.End -and $lines[$j].Trim().Length -eq 0) { $j++ }
                    if ($j -le $seg.End -and $lines[$j] -cmatch '^    VAR\s*$') {
                        $j++
                        while ($j -le $seg.End -and $lines[$j] -cnotmatch '^    \S') {
                            if ($declRx.IsMatch($lines[$j])) {
                                $sym = ConvertTo-AlSymbol -Decl $lines[$j] -Scope $procName -Origin 'Local'
                                if ($sym) { $out.Add($sym) }
                            }
                            $j++
                        }
                    }
                    # пропустить тело: от '    BEGIN' до '    END;' на том же уровне;
                    # если раньше встретился следующий заголовок — тела нет
                    $sawBegin = $false
                    while ($j -le $seg.End) {
                        if (-not $sawBegin -and $lines[$j] -cmatch '^    BEGIN\s*$') { $sawBegin = $true }
                        elseif (-not $sawBegin -and ($headRx.IsMatch($lines[$j]) -or $lines[$j] -cmatch '^    \[')) { break }
                        elseif ($sawBegin -and $lines[$j] -cmatch '^    END;\s*$') { $j++; break }
                        $j++
                    }
                    $i = $j
                    continue
                }
            }

            if ($seenProc -and $ln -cmatch '^    BEGIN\s*$') { break }  # хвост Documentation
            $i++
        }
    }

    # -- имена таблиц для Record ---------------------------------------------
    foreach ($sym in $out) {
        if ($sym.Kind -eq 'Record' -and $null -ne $sym.ObjectId) {
            $sym.TableName = Get-AlTableNameById -TableId $sym.ObjectId -SourceRoot $SourceRoot
        }
    }
    return ,$out
}

# Get-AlHeaderFuncName переехала в Lib-AlListing.ps1: имя присваивается строке
# листинга, то есть слоем ниже. Здесь она доступна по наследству — эта библиотека
# подключает Lib-AlListing.ps1 сама (см. шапку файла).

# === Get-AlStructure ========================================================
#
# Рекурсивный спуск по токенам очищенных строк. Ловушки C/AL, которые учтены:
#   - END закрывает и BEGIN, и CASE;
#   - IF..THEN без BEGIN (тело — один оператор), однострочные конструкции;
#   - вложенные CASE, ветка 'значение:' с '::' и '..' в метках;
#   - REPEAT..UNTIL <выражение> без ';' перед END.

$script:AlTokRx = [regex]'("[^"]*"|''''|[A-Za-z_][A-Za-z0-9_]*|[0-9]+(\.[0-9]+)?|::|:=|\.\.|<=|>=|<>|\S)'

function Test-AlTok {
    param([string]$Tok, [string[]]$Set)
    if ($null -eq $Tok) { return $false }
    foreach ($s in $Set) { if ($Tok -eq $s) { return $true } }
    return $false
}

function Get-AlTokPeek { param($Ctx)
    if ($Ctx.P -lt $Ctx.N) { return $Ctx.T[$Ctx.P].T }
    return $null
}

function Step-AlTok { param($Ctx)
    $tok = $Ctx.T[$Ctx.P]
    $Ctx.P++
    $Ctx.LastLine = $tok.L
    return $tok
}

function Skip-AlExpr {
    <#
    .SYNOPSIS
        Пропускает выражение до стоп-токена на нулевой глубине скобок (не съедая его).
        Возвращает стоп-токен или $null при исчерпании.
    #>
    param($Ctx, [string[]]$Stops)
    $depth = 0
    while ($Ctx.P -lt $Ctx.N) {
        $t = $Ctx.T[$Ctx.P].T
        if ($depth -eq 0 -and (Test-AlTok $t $Stops)) { return $t }
        if ($t -eq '(' -or $t -eq '[') { $depth++ }
        elseif ($t -eq ')' -or $t -eq ']') { if ($depth -gt 0) { $depth-- } }
        [void](Step-AlTok $Ctx)
    }
    return $null
}

function Read-AlStmtList {
    param($Ctx, [string[]]$Terms)
    while ($Ctx.P -lt $Ctx.N) {
        $t = Get-AlTokPeek $Ctx
        if (Test-AlTok $t $Terms) { return }
        if ($t -eq ';') { [void](Step-AlTok $Ctx); continue }
        $p0 = $Ctx.P
        Read-AlStmt $Ctx
        if ($Ctx.P -eq $p0) {
            throw ("разбор застрял на токене '{0}' (строка {1})" -f $t, $Ctx.T[$Ctx.P].L)
        }
    }
}

function Read-AlStmt {
    param($Ctx)
    $t = Get-AlTokPeek $Ctx
    if ($null -eq $t) { return }
    $up = $t.ToUpperInvariant()

    switch ($up) {
        'BEGIN' {
            [void](Step-AlTok $Ctx)
            Read-AlStmtList $Ctx @('END')
            if ((Get-AlTokPeek $Ctx) -ne 'END') { throw ("BEGIN без END (строка {0})" -f $Ctx.LastLine) }
            [void](Step-AlTok $Ctx)
            return
        }
        'IF' {
            [void](Step-AlTok $Ctx)
            $stop = Skip-AlExpr $Ctx @('THEN')
            if ($null -eq $stop) { throw ("IF без THEN (строка {0})" -f $Ctx.LastLine) }
            [void](Step-AlTok $Ctx)   # THEN
            $nx = Get-AlTokPeek $Ctx
            if (-not (Test-AlTok $nx @(';', 'ELSE', 'END', 'UNTIL'))) { Read-AlStmt $Ctx }
            if ((Get-AlTokPeek $Ctx) -eq 'ELSE') {
                [void](Step-AlTok $Ctx)
                $nx = Get-AlTokPeek $Ctx
                if ($null -ne $nx -and -not (Test-AlTok $nx @(';', 'END', 'UNTIL'))) { Read-AlStmt $Ctx }
            }
            return
        }
        'CASE' {
            [void](Step-AlTok $Ctx)
            $stop = Skip-AlExpr $Ctx @('OF')
            if ($null -eq $stop) { throw ("CASE без OF (строка {0})" -f $Ctx.LastLine) }
            [void](Step-AlTok $Ctx)   # OF
            while ($true) {
                $p = Get-AlTokPeek $Ctx
                if ($null -eq $p) { throw ("CASE без END (строка {0})" -f $Ctx.LastLine) }
                if ($p -eq ';') { [void](Step-AlTok $Ctx); continue }
                if ($p -eq 'END') { [void](Step-AlTok $Ctx); break }
                if ($p -eq 'ELSE') {
                    [void](Step-AlTok $Ctx)
                    Read-AlStmtList $Ctx @('END')
                    if ((Get-AlTokPeek $Ctx) -ne 'END') { throw ("CASE ELSE без END (строка {0})" -f $Ctx.LastLine) }
                    [void](Step-AlTok $Ctx)
                    break
                }
                # метки ветки до ':' ( '::' и '..' — отдельные токены, не мешают)
                $stop = Skip-AlExpr $Ctx @(':', 'END', 'ELSE')
                if ($stop -eq ':') {
                    [void](Step-AlTok $Ctx)
                    $nx = Get-AlTokPeek $Ctx
                    if ($null -ne $nx -and -not (Test-AlTok $nx @(';', 'END', 'ELSE'))) { Read-AlStmt $Ctx }
                }
                elseif ($null -eq $stop) { throw ("ветка CASE без ':' (строка {0})" -f $Ctx.LastLine) }
            }
            return
        }
        'REPEAT' {
            $tok = Step-AlTok $Ctx
            $node = @{ Kind = 'REPEAT'; StartLine = $tok.L; EndLine = $tok.L; Depth = $Ctx.Depth + 1 }
            $Ctx.Depth++
            Read-AlStmtList $Ctx @('UNTIL')
            if ((Get-AlTokPeek $Ctx) -ne 'UNTIL') { throw ("REPEAT без UNTIL (строка {0})" -f $node.StartLine) }
            [void](Step-AlTok $Ctx)   # UNTIL
            $stop = Skip-AlExpr $Ctx @(';', 'END', 'ELSE', 'UNTIL', ':')
            if ($stop -eq ';') { [void](Step-AlTok $Ctx) }
            $Ctx.Depth--
            $node.EndLine = $Ctx.LastLine
            $Ctx.Loops.Add($node)
            return
        }
        'FOR' {
            $tok = Step-AlTok $Ctx
            $node = @{ Kind = 'FOR'; StartLine = $tok.L; EndLine = $tok.L; Depth = $Ctx.Depth + 1 }
            $stop = Skip-AlExpr $Ctx @('DO')
            if ($null -eq $stop) { throw ("FOR без DO (строка {0})" -f $node.StartLine) }
            [void](Step-AlTok $Ctx)   # DO
            $Ctx.Depth++
            $nx = Get-AlTokPeek $Ctx
            if ($null -ne $nx -and -not (Test-AlTok $nx @(';', 'END', 'ELSE', 'UNTIL'))) { Read-AlStmt $Ctx }
            $Ctx.Depth--
            $node.EndLine = $Ctx.LastLine
            $Ctx.Loops.Add($node)
            return
        }
        'WHILE' {
            $tok = Step-AlTok $Ctx
            $node = @{ Kind = 'WHILE'; StartLine = $tok.L; EndLine = $tok.L; Depth = $Ctx.Depth + 1 }
            $stop = Skip-AlExpr $Ctx @('DO')
            if ($null -eq $stop) { throw ("WHILE без DO (строка {0})" -f $node.StartLine) }
            [void](Step-AlTok $Ctx)   # DO
            $Ctx.Depth++
            $nx = Get-AlTokPeek $Ctx
            if ($null -ne $nx -and -not (Test-AlTok $nx @(';', 'END', 'ELSE', 'UNTIL'))) { Read-AlStmt $Ctx }
            $Ctx.Depth--
            $node.EndLine = $Ctx.LastLine
            $Ctx.Loops.Add($node)
            return
        }
        'WITH' {
            $tok = Step-AlTok $Ctx
            $node = @{ Target = ''; StartLine = $tok.L; EndLine = $tok.L }
            $sb = New-Object System.Text.StringBuilder
            while ($true) {
                $p = Get-AlTokPeek $Ctx
                if ($null -eq $p) { throw ("WITH без DO (строка {0})" -f $node.StartLine) }
                if ($p -eq 'DO') { [void](Step-AlTok $Ctx); break }
                [void]$sb.Append((Step-AlTok $Ctx).T)
            }
            $node.Target = $sb.ToString()
            $nx = Get-AlTokPeek $Ctx
            if ($null -ne $nx -and -not (Test-AlTok $nx @(';', 'END', 'ELSE', 'UNTIL'))) { Read-AlStmt $Ctx }
            $node.EndLine = $Ctx.LastLine
            $Ctx.Withs.Add($node)
            return
        }
        default {
            # простой оператор: до ';' (съедаем) или до закрывающего слова (не съедаем)
            if (Test-AlTok $t @('END', 'ELSE', 'UNTIL')) {
                throw ("неожиданный токен '{0}' (строка {1})" -f $t, $Ctx.T[$Ctx.P].L)
            }
            $stop = Skip-AlExpr $Ctx @(';', 'END', 'ELSE', 'UNTIL')
            if ($stop -eq ';') { [void](Step-AlTok $Ctx) }
            return
        }
    }
}

function Get-AlStructure {
    <#
    .SYNOPSIS
        Структура кода объекта: функции, циклы, WITH; глубина циклов на каждую строку.

    .DESCRIPTION
        Возвращает объект с двумя наборами:
          Functions — на каждую функцию: Name, HeaderLine, FirstLine, LastLine,
                      Reliable (разбор сошёлся), Reason (почему нет), LoopCount,
                      Loops (Kind, StartLine, EndLine, Depth), Withs (Target, StartLine, EndLine).
          Lines     — на каждую строку кода: LineNo, Function, LoopDepth,
                      LoopKind/LoopStart/LoopEnd (ближайший объемлющий цикл), WithVar.
        При Reliable=false у функции циклы могут быть неполными — правила должны
        понижать уверенность подсказок по такой функции.
    #>
    param(
        [Parameter(Mandatory)][object]$Listing,
        [object]$Lexed
    )
    $rows = ConvertTo-AlRows $Listing
    if ($Lexed) { $lex = ConvertTo-AlRows $Lexed }
    else        { $lex = ConvertTo-AlRows (Get-AlLexed -Listing $rows) }
    $funcMap = ConvertTo-AlRows (Get-AlFunctionMap -Listing $rows)

    $functions = New-Object System.Collections.Generic.List[object]
    $lineRecs  = New-Object System.Collections.Generic.List[object]

    foreach ($fn in $funcMap) {
        # Имя берётся из карты как есть: Get-AlListing присваивает его через
        # Get-AlHeaderFuncName, и переопределять больше нечего. Раньше здесь стояло
        # повторное чтение заголовка по индексу $rows[$fn.HeaderLine - 1] — оно молча
        # держалось на том, что позиция в массиве равна номеру строки минус один.
        $fnName = $fn.Name

        # токены тела функции (только строки Kind=Code, по очищенному тексту)
        $tokens = New-Object System.Collections.Generic.List[object]
        for ($l = $fn.FirstLine; $l -le $fn.LastLine; $l++) {
            $row = $lex[$l - 1]
            if ($row.Kind -ne 'Code' -or $row.Clean.Length -eq 0) { continue }
            foreach ($m in $script:AlTokRx.Matches($row.Clean)) {
                $v = $m.Value
                if ($v.Trim().Length -eq 0) { continue }
                $tokens.Add(@{ T = $v; L = $l })
            }
        }

        $ctx = @{
            T = $tokens; N = $tokens.Count; P = 0
            Depth = 0; LastLine = $fn.HeaderLine
            Loops = New-Object System.Collections.Generic.List[object]
            Withs = New-Object System.Collections.Generic.List[object]
        }
        $reliable = $true; $reason = ''
        try {
            Read-AlStmtList $ctx @()
            if ($ctx.P -lt $ctx.N) {
                $reliable = $false
                $reason = ("не все токены разобраны, остановка на '{0}' (строка {1})" -f `
                    $ctx.T[$ctx.P].T, $ctx.T[$ctx.P].L)
            }
        }
        catch {
            $reliable = $false
            $reason = $_.Exception.Message
        }

        $loops = @($ctx.Loops | ForEach-Object {
            [pscustomobject]@{ Kind = $_.Kind; StartLine = $_.StartLine; EndLine = $_.EndLine; Depth = $_.Depth }
        })
        $withs = @($ctx.Withs | ForEach-Object {
            [pscustomobject]@{ Target = $_.Target; StartLine = $_.StartLine; EndLine = $_.EndLine }
        })

        $functions.Add([pscustomobject]@{
            Name       = $fnName
            HeaderLine = $fn.HeaderLine
            FirstLine  = $fn.FirstLine
            LastLine   = $fn.LastLine
            Reliable   = $reliable
            Reason     = $reason
            LoopCount  = $loops.Count
            Loops      = $loops
            Withs      = $withs
        })

        # построчная сводка: глубина и ближайший объемлющий цикл, цель WITH
        for ($l = $fn.FirstLine; $l -le $fn.LastLine; $l++) {
            $row = $lex[$l - 1]
            if ($row.Kind -ne 'Code') { continue }
            $depth = 0; $inner = $null
            foreach ($lp in $loops) {
                if ($l -ge $lp.StartLine -and $l -le $lp.EndLine) {
                    $depth++
                    if ($null -eq $inner -or $lp.StartLine -gt $inner.StartLine) { $inner = $lp }
                }
            }
            $withVar = ''
            $innerW = $null
            foreach ($w in $withs) {
                if ($l -ge $w.StartLine -and $l -le $w.EndLine) {
                    if ($null -eq $innerW -or $w.StartLine -gt $innerW.StartLine) { $innerW = $w }
                }
            }
            if ($null -ne $innerW) { $withVar = $innerW.Target }

            $loopKind = ''; $loopStart = 0; $loopEnd = 0
            if ($null -ne $inner) { $loopKind = $inner.Kind; $loopStart = $inner.StartLine; $loopEnd = $inner.EndLine }
            $lineRecs.Add([pscustomobject]@{
                LineNo    = $l
                Function  = $fnName
                LoopDepth = $depth
                LoopKind  = $loopKind
                LoopStart = $loopStart
                LoopEnd   = $loopEnd
                WithVar   = $withVar
            })
        }
    }

    return [pscustomobject]@{
        Functions = $functions.ToArray()
        Lines     = $lineRecs.ToArray()
    }
}

# === Get-AlFilterChains =====================================================

$script:AlFilterOpsRx = [regex]('(?<![\w".])(?:(?<var>"[^"]+"|[A-Za-z_][A-Za-z0-9_]*)\s*\.\s*)?' +
    '(?<op>SETRANGE|SETFILTER|SETCURRENTKEY|RESET|FINDSET|FINDFIRST|FINDLAST|FIND|COUNT|ISEMPTY|CALCSUMS|DELETEALL|MODIFYALL)(?![\w])')

function Get-AlFilterChains {
    <#
    .SYNOPSIS
        Цепочки «строки фильтров -> потребитель» по Record-переменным внутри функций.

    .DESCRIPTION
        Для каждой Record-переменной копится последовательность SETRANGE/SETFILTER/
        SETCURRENTKEY/RESET; ближайший ниже по коду вызов FINDSET/FINDFIRST/FINDLAST/
        FIND/COUNT/ISEMPTY/CALCSUMS/DELETEALL/MODIFYALL той же переменной замыкает цепочку.
        Подсказка вешается на потребителя (у него метрики), состав фильтра — строки выше.

        На каждый потребитель возвращается:
          FilterLines       — фильтры, поставленные ПОСЛЕ предыдущего потребителя;
          ActiveFilterLines — всё действующее с последнего RESET (фильтры в C/AL липкие).
        RESET сбрасывает накопленное и сам входит в новую цепочку.

        Вызовы без 'Переменная.' приписываются цели объемлющего WITH, а вне WITH — Rec
        (для таблиц Rec/xRec разрешаются в саму таблицу, передайте -ObjectType/-ObjectId).
        Если переданы -Symbols, вызовы по переменным НЕ типа Record отбрасываются
        (защита от DotNet .COUNT и т.п.), а у цепочки заполняются TableNo/Temporary;
        переменная, которой нет в символах, остаётся с Resolved=$false.
    #>
    param(
        [Parameter(Mandatory)][object]$Listing,
        [object]$Lexed,
        [object]$Structure,
        [object]$Symbols,
        [int]$ObjectType = 0,
        [int]$ObjectId   = 0
    )
    $rows = ConvertTo-AlRows $Listing
    if ($Lexed) { $lex = ConvertTo-AlRows $Lexed }
    else        { $lex = ConvertTo-AlRows (Get-AlLexed -Listing $rows) }
    if (-not $Structure) { $Structure = Get-AlStructure -Listing $rows -Lexed $lex }
    $syms = ConvertTo-AlRows $Symbols

    # WithVar по номеру строки
    $withByLine = @{}
    foreach ($lr in $Structure.Lines) {
        if ($lr.WithVar) { $withByLine[$lr.LineNo] = $lr.WithVar }
    }

    # индекс символов: 'область|имя' -> символ (локальные и параметры), '|имя' -> глобальный
    $symIdx = $null
    if ($syms.Count -gt 0) {
        $symIdx = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($s in $syms) { $symIdx[($s.Scope + '|' + $s.Name)] = $s }
    }

    $chains = New-Object System.Collections.Generic.List[object]
    $filterOps = @('SETRANGE', 'SETFILTER', 'SETCURRENTKEY', 'RESET')

    foreach ($fn in $Structure.Functions) {
        $state = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)

        for ($l = $fn.FirstLine; $l -le $fn.LastLine; $l++) {
            $row = $lex[$l - 1]
            if ($row.Kind -ne 'Code' -or $row.Clean.Length -eq 0) { continue }

            foreach ($m in $script:AlFilterOpsRx.Matches($row.Clean)) {
                $op  = $m.Groups['op'].Value
                $var = ''
                $viaWith = $false
                if ($m.Groups['var'].Success) {
                    $var = $m.Groups['var'].Value.Trim('"')
                }
                elseif ($withByLine.ContainsKey($l)) {
                    $var = $withByLine[$l].Trim('"'); $viaWith = $true
                }
                else {
                    $var = 'Rec'
                }

                # разрешение типа переменной
                $resolved = $false; $tableNo = 0; $tableName = ''; $temporary = $false
                if (($var -eq 'Rec' -or $var -eq 'xRec') -and $ObjectType -eq 1 -and $ObjectId -gt 0) {
                    $resolved = $true; $tableNo = $ObjectId
                    $tableName = Get-AlTableNameById -TableId $ObjectId
                }
                elseif ($null -ne $symIdx) {
                    $sym = $null
                    if (-not $symIdx.TryGetValue(($fn.Name + '|' + $var), [ref]$sym)) {
                        [void]$symIdx.TryGetValue(('|' + $var), [ref]$sym)
                    }
                    if ($null -ne $sym) {
                        if ($sym.Kind -ne 'Record') { continue }   # не Record — не наш вызов
                        $resolved = $true; $tableNo = $sym.ObjectId
                        $tableName = $sym.TableName; $temporary = $sym.Temporary
                    }
                }

                if (-not $state.ContainsKey($var)) {
                    $state[$var] = @{
                        Active = New-Object System.Collections.Generic.List[object]
                        Fresh  = New-Object System.Collections.Generic.List[object]
                    }
                }
                $st = $state[$var]

                if ($filterOps -contains $op) {
                    if ($op -eq 'RESET') { $st.Active.Clear(); $st.Fresh.Clear() }
                    $item = [pscustomobject]@{ LineNo = $l; Op = $op }
                    $st.Active.Add($item)
                    $st.Fresh.Add($item)
                }
                else {
                    $chains.Add([pscustomobject]@{
                        Function          = $fn.Name
                        Variable          = $var
                        Resolved          = $resolved
                        TableNo           = $tableNo
                        TableName         = $tableName
                        Temporary         = $temporary
                        ViaWith           = $viaWith
                        ConsumerLine      = $l
                        ConsumerOp        = $op
                        FilterLines       = $st.Fresh.ToArray()
                        ActiveFilterLines = $st.Active.ToArray()
                    })
                    $st.Fresh = New-Object System.Collections.Generic.List[object]
                }
            }
        }
    }
    return ,$chains
}
