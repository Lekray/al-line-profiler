#Requires -Version 5.1
<#
.SYNOPSIS
    Чтение оглавления .fob и заголовков текстового экспорта C/SIDE. Подключается
    через dot-sourcing.

.DESCRIPTION
    Контейнер .fob начинается с ОГЛАВЛЕНИЯ открытым текстом: на каждый объект две
    строки — «тип, номер, имя, дата, время» и «размер, список версий». Дальше идёт
    двоичное тело, и читать его незачем.

    Что из этого следует и чего НЕ следует. Оглавление позволяет сверить состав:
    те же ли объекты, те же ли имена, даты и списки версий. Но доказать, что .fob
    собран ИМЕННО ИЗ этого текста, оглавление не может: дата в OBJECT-PROPERTIES
    ставится руками по соглашению проекта (время у нас всегда 12:00:00), и правка
    кода в тот же день оглавление не меняет вовсе. Проверено на живом расхождении:
    .fob отставал от текста на восемь коммитов, а состав совпадал полностью.

    Поэтому происхождение объявляется отдельно — манифестом SHA256SUMS.txt рядом с
    файлами. Он фиксирует хэш текста, С КОТОРОГО снят .fob. Это декларация, а не
    доказательство, но декларация проверяемая: любая последующая правка текста её
    ломает, и сторож говорит, что .fob пора перевыгрузить.

    Подключение:  . (Join-Path $PSScriptRoot 'Lib-CSideFob.ps1')
#>

function Get-CSideFobDirectory {
    <#
    .SYNOPSIS
        Оглавление .fob: тип, номер, имя, дата, время, размер, список версий.

    .DESCRIPTION
        Заголовок читается как ANSI (cp1251): имена объектов проекта латинские, а
        для кириллических имён сверка по имени всё равно даст расхождение, и это
        честнее молчания. Разбор останавливается на первой строке, которая под
        формат оглавления не подходит, — дальше начинается двоичное тело.

    .PARAMETER Path
        Путь к .fob.

    .PARAMETER HeadBytes
        Сколько байт от начала считать заголовком. Хватает 160 байт на объект;
        по умолчанию — с запасом на пару сотен объектов.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Path,
        [int] $HeadBytes = 65536
    )

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $take  = [math]::Min($HeadBytes, $bytes.Length)
    $enc   = [System.Text.Encoding]::GetEncoding(1251)
    $head  = $enc.GetString($bytes, 0, $take)
    $lines = $head -split "`r`n"

    $rxObj  = [regex]'^([A-Za-z]+)\s+(\d+)\s+(.*?)\s+(\d{2}\.\d{2}\.\d{2})\s+(\d{2}:\d{2}:\d{2})\s*$'
    $rxSize = [regex]'^\s+(\d+)\s+(\S*)\s*$'

    $out = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt ($lines.Count - 1)) {
        $m = $rxObj.Match($lines[$i])
        if (-not $m.Success) { break }
        $size = 0; $vl = ''
        $m2 = $rxSize.Match($lines[$i + 1])
        if ($m2.Success) { $size = [int]$m2.Groups[1].Value; $vl = $m2.Groups[2].Value }
        [void]$out.Add([pscustomobject]@{
            Type        = $m.Groups[1].Value
            Id          = [int]$m.Groups[2].Value
            Name        = $m.Groups[3].Value
            Date        = $m.Groups[4].Value
            Time        = $m.Groups[5].Value
            Size        = $size
            VersionList = $vl
        })
        $i += 2
    }
    return $out.ToArray()
}

function Get-CSideTextDirectory {
    <#
    .SYNOPSIS
        То же по текстовому экспорту: заголовок OBJECT плюс блок OBJECT-PROPERTIES.

    .PARAMETER Path
        Путь к .txt (UTF-8 без BOM, CRLF).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Path)

    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
    $rx = [regex]::new(
        '^OBJECT\s+(\S+)\s+(\d+)\s+(.*?)\r?\n\{\r?\n\s*OBJECT-PROPERTIES\r?\n\s*\{(.*?)\r?\n\s*\}',
        ([System.Text.RegularExpressions.RegexOptions]'Multiline,Singleline'))

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($m in $rx.Matches($text)) {
        $body = $m.Groups[4].Value
        $get = {
            param([string] $Key)
            $mm = [regex]::Match($body, [regex]::Escape($Key) + '\s*=([^;]*);')
            if ($mm.Success) { return $mm.Groups[1].Value.Trim() }
            return ''
        }
        [void]$out.Add([pscustomobject]@{
            Type        = $m.Groups[1].Value
            Id          = [int]$m.Groups[2].Value
            Name        = $m.Groups[3].Value.Trim()
            Date        = (& $get 'Date')
            Time        = (& $get 'Time')
            Size        = 0
            VersionList = (& $get 'Version List')
        })
    }
    return $out.ToArray()
}

function Compare-CSideDirectory {
    <#
    .SYNOPSIS
        Сверка состава .fob и текстового экспорта. Возвращает список расхождений
        строками; пустой список — состав совпал.

    .DESCRIPTION
        Сверяются состав объектов, имя, дата, время и список версий. Совпадение
        состава НЕ означает, что .fob собран из этого текста, — см. шапку файла.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]] $Fob,
        [Parameter(Mandatory)][object[]] $Text
    )

    $fx = @{}; foreach ($o in $Fob)  { $fx['{0}/{1}' -f $o.Type, $o.Id] = $o }
    $tx = @{}; foreach ($o in $Text) { $tx['{0}/{1}' -f $o.Type, $o.Id] = $o }

    $keys = @(@($fx.Keys) + @($tx.Keys) | Sort-Object -Unique)
    $diff = New-Object System.Collections.Generic.List[string]
    foreach ($k in $keys) {
        $f = $fx[$k]; $t = $tx[$k]
        if ($null -eq $t) { [void]$diff.Add(('{0}: есть в .fob, нет в тексте' -f $k)); continue }
        if ($null -eq $f) { [void]$diff.Add(('{0}: есть в тексте, нет в .fob' -f $k)); continue }
        if ($f.Name -ne $t.Name) {
            [void]$diff.Add(('{0}: имя «{1}» в .fob против «{2}» в тексте' -f $k, $f.Name, $t.Name))
        }
        if ($f.Date -ne $t.Date -or $f.Time -ne $t.Time) {
            [void]$diff.Add(('{0}: правка {1} {2} в .fob против {3} {4} в тексте' -f
                             $k, $f.Date, $f.Time, $t.Date, $t.Time))
        }
        if ($f.VersionList -ne $t.VersionList) {
            [void]$diff.Add(('{0}: список версий «{1}» в .fob против «{2}» в тексте' -f
                             $k, $f.VersionList, $t.VersionList))
        }
    }
    return $diff.ToArray()
}

function Get-CSideSha256 {
    <#
    .SYNOPSIS
        SHA-256 файла строчными шестнадцатеричными, как печатает sha256sum.
    #>
    param([Parameter(Mandatory)][string] $Path)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        try { return ([BitConverter]::ToString($sha.ComputeHash($fs))).Replace('-', '').ToLowerInvariant() }
        finally { $fs.Dispose() }
    }
    finally { $sha.Dispose() }
}

function Read-CSideSumsFile {
    <#
    .SYNOPSIS
        Манифест sha256sum: «хэш<пробел><пробел|*>имя». Возвращает хэш-таблицу
        «имя файла -> хэш».

    .DESCRIPTION
        Читается и вариант со звёздочкой (двоичный режим sha256sum), и без неё.
        Имя берётся как есть, без пути: манифест лежит рядом со своими файлами.
    #>
    param([Parameter(Mandatory)][string] $Path)
    $map = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $map }
    foreach ($line in [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)) {
        $m = [regex]::Match($line.Trim(), '^([0-9a-fA-F]{64})\s+\*?(.+)$')
        if ($m.Success) { $map[$m.Groups[2].Value.Trim()] = $m.Groups[1].Value.ToLowerInvariant() }
    }
    return $map
}
