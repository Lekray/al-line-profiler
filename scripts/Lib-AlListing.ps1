<#
.SYNOPSIS
    Библиотека разбора дампа исходников C/AL. Подключается через dot-sourcing.

.DESCRIPTION
    Дамп, который делает Dump-AlSource.ps1, — это код объекта в том виде, в каком его
    хранит платформа: заголовок функции или триггера начинается с нулевой колонки, тело
    отбито двумя пробелами. Объявления переменных в этот дамп не входят, то есть почти
    каждая непустая строка — либо заголовок, либо комментарий, либо исполняемый код.

    Нумерация строк здесь платформенная: она совпадает с номерами в событиях трассировки
    и НЕ совпадает с нумерацией текстового экспорта C/SIDE (каталога baseline).

    Подключение:  . (Join-Path $PSScriptRoot 'Lib-AlListing.ps1')
#>
#Requires -Version 5.1

function Get-AlTypeNames {
    <#
    .SYNOPSIS
        Типы объектов NAV — как в dbo.[Object].[Type] и в событиях трассировки.

    .DESCRIPTION
        Таблица заполняется по факту обращения, а не присваиванием на уровне файла,
        и это не стилистика. У функции, подключённой через dot-source, «$script:»
        разрешается в область ВЫЗЫВАЮЩЕГО скрипта, а не того файла, где функция
        объявлена. Потребитель, которому функции достались по наследству (сам
        dot-source не делал — например, потому что вызывающий скрипт подключил
        библиотеку раньше), увидит переменную пустой, и падение случится у него.
    #>
    if (-not $script:AlTypeNames) {
        $script:AlTypeNames = @{
            1 = 'Table'; 2 = 'Form';      3 = 'Report'; 4 = 'Dataport'; 5 = 'Codeunit'
            6 = 'XMLport'; 7 = 'MenuSuite'; 8 = 'Page'; 9 = 'Query';   10 = 'System'
        }
    }
    return $script:AlTypeNames
}

function Get-AlTypeName {
    param([int]$TypeId)
    $names = Get-AlTypeNames
    if ($names.ContainsKey($TypeId)) { return $names[$TypeId] }
    return "Type$TypeId"
}

function Get-AlSourceRoot {
    <#
    .SYNOPSIS
        Каталог дампа исходников (по умолчанию <корень репозитория>\.alsrc).
    #>
    param([string]$Root)
    if ($Root) { return $Root }
    $repoRoot = Split-Path -Parent $PSScriptRoot
    return (Join-Path $repoRoot '.alsrc')
}

function Get-AlObjectInfo {
    <#
    .SYNOPSIS
        Строка из index.tsv по типу и номеру объекта: имя, дата, время, хэш.
    #>
    param(
        [Parameter(Mandatory)][int]$ObjectType,
        [Parameter(Mandatory)][int]$ObjectId,
        [string]$SourceRoot
    )
    $root = Get-AlSourceRoot $SourceRoot
    $idx  = Join-Path $root 'index.tsv'
    if (-not (Test-Path -LiteralPath $idx)) {
        throw "Не найден индекс дампа: $idx. Сначала выполните Dump-AlSource.ps1."
    }
    $lines = [System.IO.File]::ReadAllLines($idx, [System.Text.Encoding]::UTF8)
    for ($i = 1; $i -lt $lines.Length; $i++) {
        $c = $lines[$i] -split "`t"
        if ($c.Length -ge 11 -and [int]$c[0] -eq $ObjectType -and [int]$c[2] -eq $ObjectId) {
            return [pscustomobject]@{
                ObjectType  = $ObjectType
                TypeName    = $c[1]
                ObjectId    = $ObjectId
                Name        = $c[3]
                Lines       = [int]$c[4]
                Bytes       = [int]$c[5]
                Compiled    = ([int]$c[6] -ne 0)
                Date        = $c[7]
                Time        = $c[8]
                VersionList = $c[9]
                Hash        = $c[10]
            }
        }
    }
    return $null
}

function Get-AlListing {
    <#
    .SYNOPSIS
        Разбирает исходник объекта в набор строк листинга.

    .DESCRIPTION
        Возвращает по объекту, на каждую строку исходника:
            LineNo       — платформенный номер строки (1-based)
            Kind         — Function | Comment | Code | Empty
            FunctionName — функция, которой принадлежит строка ('' до первого заголовка)
            Indent       — уровень для дерева: 1 заголовок функции, 2 строка кода
            Text         — исходный текст строки как есть
    #>
    param(
        [Parameter(Mandatory)][int]$ObjectType,
        [Parameter(Mandatory)][int]$ObjectId,
        [string]$SourceRoot
    )
    $root = Get-AlSourceRoot $SourceRoot
    $path = Join-Path $root ("{0}_{1}.al" -f $ObjectType, $ObjectId)
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Нет исходника $path. Объект не выгружен или не скомпилирован в базе."
    }

    $text  = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $lines = $text -split "`r`n"

    $result   = New-Object System.Collections.Generic.List[object]
    $currFunc = ''
    $lineNo   = 0

    foreach ($raw in $lines) {
        $lineNo++
        $trimmed = $raw.Trim()

        if ($trimmed.Length -eq 0) {
            $kind = 'Empty'; $indent = 2
        }
        elseif (-not [char]::IsWhiteSpace($raw[0])) {
            # нулевая колонка занята — это заголовок функции или триггера
            $kind = 'Function'; $indent = 1
            $paren = $raw.IndexOf('(')
            if ($paren -gt 0) { $currFunc = $raw.Substring(0, $paren) } else { $currFunc = $trimmed }
        }
        elseif ($trimmed.StartsWith('//')) {
            $kind = 'Comment'; $indent = 2
        }
        else {
            $kind = 'Code'; $indent = 2
        }

        $result.Add([pscustomobject]@{
            LineNo       = $lineNo
            Kind         = $kind
            FunctionName = $currFunc
            Indent       = $indent
            Text         = $raw
        })
    }
    # КОНТРАКТ: возвращаем [object[]], вызывающая сторона ОБЯЗАНА обернуть в @().
    # Почему именно так:
    #   * «return ,$arr» ломает @(...) — обёртка считает весь массив одним элементом;
    #   * возврат List[object] ломается иначе: на этой сборке WinPS 5.1 выражение
    #     @(<List[object]>) внутри функции падает с «Argument types do not match»
    #     (PSToObjectArrayBinder). Поэтому именно .ToArray() и именно без запятой.
    return $result.ToArray()
}

function Get-AlFunctionMap {
    <#
    .SYNOPSIS
        Карта функций объекта: имя, строка заголовка, границы тела.

    .DESCRIPTION
        Нужна, чтобы переводить «функция + относительный номер строки» (так номер строки
        приходит в стеке события о долгом SQL) в абсолютный номер строки листинга, а также
        чтобы считать сводки по функциям.
    #>
    param([Parameter(Mandatory)][object[]]$Listing)

    # Один проход: листинг упорядочен по номерам строк, поэтому границы функций
    # и счётчики набираются на лету. Перебор Where-Object внутри цикла давал бы
    # квадратичную сложность, а линтеру ходить по всем 7954 объектам базы.
    $map     = New-Object System.Collections.Generic.List[object]
    $arr     = @($Listing)
    $curr    = $null
    $codeCnt = 0
    $totCnt  = 0

    foreach ($l in $arr) {
        if ($l.Kind -eq 'Function') {
            if ($curr) {
                $curr.LastLine   = $l.LineNo - 1
                $curr.CodeLines  = $codeCnt
                $curr.TotalLines = $totCnt
                $map.Add($curr)
            }
            $curr = [pscustomobject]@{
                Name       = $l.FunctionName
                HeaderLine = $l.LineNo
                FirstLine  = $l.LineNo + 1
                LastLine   = $l.LineNo
                CodeLines  = 0
                TotalLines = 0
                Signature  = $l.Text
            }
            $codeCnt = 0
            $totCnt  = 0
        }
        elseif ($curr) {
            $totCnt++
            if ($l.Kind -eq 'Code') { $codeCnt++ }
        }
    }
    if ($curr) {
        $curr.LastLine   = $arr[$arr.Count - 1].LineNo
        $curr.CodeLines  = $codeCnt
        $curr.TotalLines = $totCnt
        $map.Add($curr)
    }

    # См. контракт в Get-AlListing: [object[]] без запятой, вызов через @().
    return $map.ToArray()
}

function ConvertTo-HtmlText {
    <#
    .SYNOPSIS
        Экранирование текста для вставки в HTML.
    #>
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
}
