#requires -Version 7
<#
.SYNOPSIS
    Объекты, изменённые после последней загрузки на объект. Пакетом или голым файлом.

.DESCRIPTION
    На объект уезжают только изменённые объекты, а не весь текст: полный экспорт - это
    444 КБ одного cp866, а правится обычно один-два объекта. Текстовый экспорт режется
    по строкам OBJECT и сравнивается с тем, что уже загружено; C/SIDE принимает
    частичный экспорт и заменяет ровно те объекты, что в нём есть.

    БАЗА СРАВНЕНИЯ - delivered.txt в корне репозитория: по объекту его SHA-256 на момент
    последней ЗАГРУЗКИ. Не сборки пакета, а именно загрузки - сборка ещё ничего не
    меняет на объекте. Обновляется ключом -Record, и только после того, как загрузка
    действительно прошла.

    Ключ -Previous заставляет сравнивать с конкретным пакетом вместо delivered.txt -
    нужен, когда база сравнения ещё не заведена или уехало что-то другое.

.PARAMETER ObjectsOnly
    Отдать ГОЛЫЕ объекты, без zip, README и манифеста - когда нужен просто файл, который
    импортируют. Файл на отправку кладётся в out\send и лежит там ОДИН: каталог чистится
    при каждой сборке, поэтому правило без оговорок - отправлять всё, что в нём лежит.
    Тот же текст в UTF-8 остаётся в out\read, для чтения глазами, и рядом с отправляемым
    не лежит НИКОГДА: пока два похожих имени соседствовали, однажды уехало не то.

.PARAMETER Previous
    Пакет, с которым сравнивать, вместо delivered.txt.

.PARAMETER Record
    Записать delivered.txt по текущему тексту: считать, что эти объекты загружены.

.PARAMETER PartChars
    Знаков полезной нагрузки в одном письме. По умолчанию 448000 - около 450 КБ тела,
    столько уходит одним письмом. Умолчание самого упаковщика (44800) резало бы пакет
    на пяток писем без нужды, а каждое лишнее письмо - это ещё один шанс отправить не
    то и собрать не в том порядке. Уменьшать, если почта принимающей стороны строже.

.PARAMETER OutDir
    Куда положить. По умолчанию out\ в корне репозитория.

.EXAMPLE
    pwsh scripts/New-DeltaPackage.ps1
    Пакет: архив с объектами, APPLY.txt, CHANGES.txt и манифестом - и сразу тело письма,
    которым он уезжает. Почта контура вложений не принимает, поэтому отправляется НЕ
    архив, а его текст: письма ложатся в out\send, сам архив - в out\read для сверки.

.EXAMPLE
    pwsh scripts/New-DeltaPackage.ps1 -ObjectsOnly
    Файл с изменёнными объектами - то, что уезжает на объект.

.EXAMPLE
    pwsh scripts/New-DeltaPackage.ps1 -Record
    После загрузки: теперь текущие объекты считаются загруженными.
#>
[CmdletBinding()]
param(
    [string] $Previous,
    [string] $OutDir,
    [switch] $ObjectsOnly,
    [switch] $Record,
    [int]    $PartChars = 448000
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$ledger = Join-Path $repo 'delivered.txt'

function Step([string]$m) { Write-Host "  $m" }

# Объект начинается со строки OBJECT в первой позиции и тянется до следующей такой.
# Иначе резать нечем: в текстовом экспорте C/SIDE нет ни разделителя, ни оглавления.
function Split-CSideObjects {
    param([string]$Text)
    $map = [ordered]@{}
    $starts = [regex]::Matches($Text, '(?m)^OBJECT ')
    for ($i = 0; $i -lt $starts.Count; $i++) {
        $from = $starts[$i].Index
        $to = if ($i + 1 -lt $starts.Count) { $starts[$i + 1].Index } else { $Text.Length }
        $block = $Text.Substring($from, $to - $from)
        $head = ($block -split "`r`n", 2)[0]
        $key = if ($head -match '^OBJECT\s+(\S+)\s+(\d+)') { $Matches[1] + ' ' + $Matches[2] } else { $head }
        $map[$key] = $block
    }
    return $map
}

function Get-TextHash {
    param([string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text))).Replace('-', '').ToLower() }
    finally { $sha.Dispose() }
}

$curFile = Join-Path $repo 'LineProfiler.txt'
$curText = [IO.File]::ReadAllText($curFile, [Text.UTF8Encoding]::new($false))
$cur = Split-CSideObjects $curText

# --- запись базы сравнения ---------------------------------------------------
if ($Record) {
    $src = $cur
    $what = 'текущий LineProfiler.txt'
    if ($Previous) {
        # Заведение базы задним числом: считать загруженным не текущий текст, а тот,
        # что уехал пакетом. Нужно ровно один раз - когда база заводится впервые.
        $z = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Previous).Path)
        $prevText = ''
        try {
            foreach ($e in $z.Entries) {
                if ($e.FullName -ne 'objects/LineProfiler.txt') { continue }
                $ms = New-Object IO.MemoryStream
                $s = $e.Open(); $s.CopyTo($ms); $s.Close()
                $prevText = [Text.Encoding]::UTF8.GetString($ms.ToArray()); $ms.Dispose()
            }
        } finally { $z.Dispose() }
        if (-not $prevText) { throw "в пакете нет objects\LineProfiler.txt" }
        $src = Split-CSideObjects $prevText
        $what = (Split-Path $Previous -Leaf)
    }
    $l = @(
        '# Что загружено на объект: по объекту - SHA-256 его текста в LineProfiler.txt.',
        '# Отсюда считается, какие объекты менялись ПОСЛЕ загрузки. Обновлять только',
        '# после того, как загрузка действительно прошла: pwsh scripts/New-DeltaPackage.ps1 -Record',
        ''
    )
    $l += ('# Источник записи: ' + $what)
    $l += ''
    foreach ($k in $src.Keys) { $l += ('{0,-20} {1}' -f $k, (Get-TextHash $src[$k])) }
    [IO.File]::WriteAllLines($ledger, $l, [Text.UTF8Encoding]::new($false))
    Write-Host ("delivered.txt записан по " + $what + ": объектов " + $src.Count) -ForegroundColor Green
    exit 0
}

Write-Host 'Изменённые объекты' -ForegroundColor Cyan

# --- база сравнения ----------------------------------------------------------
$baseHash = @{}
$baseName = ''
$prevHash = @{}
if ($Previous) {
    if (-not (Test-Path -LiteralPath $Previous)) { throw "не найден пакет: $Previous" }
    $z = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Previous).Path)
    $prevText = ''
    try {
        foreach ($e in $z.Entries) {
            if ($e.FullName -like '*/') { continue }
            $ms = New-Object IO.MemoryStream
            $s = $e.Open(); $s.CopyTo($ms); $s.Close()
            $bytes = $ms.ToArray(); $ms.Dispose()
            $key = $e.FullName -replace '/', '\'
            $sha = [Security.Cryptography.SHA256]::Create()
            $prevHash[$key] = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '')
            $sha.Dispose()
            if ($key -eq 'objects\LineProfiler.txt') { $prevText = [Text.Encoding]::UTF8.GetString($bytes) }
        }
    } finally { $z.Dispose() }
    if (-not $prevText) { throw "в пакете нет objects\LineProfiler.txt" }
    $prev = Split-CSideObjects $prevText
    foreach ($k in $prev.Keys) { $baseHash[$k] = Get-TextHash $prev[$k] }
    $baseName = Split-Path $Previous -Leaf
} elseif (Test-Path $ledger) {
    foreach ($line in (Get-Content $ledger)) {
        if ($line -match '^\s*#' -or $line.Trim() -eq '') { continue }
        if ($line -match '^(\S+\s+\d+)\s+([0-9a-f]{64})$') { $baseHash[$Matches[1]] = $Matches[2] }
    }
    $baseName = 'delivered.txt'
} else {
    throw ("нет базы сравнения: ни -Previous, ни delivered.txt. " +
           "Заведите её после первой загрузки: pwsh scripts/New-DeltaPackage.ps1 -Record")
}
Step ("сравнение с: $baseName, объектов в базе " + $baseHash.Count)

$take = @()
$added = @()
foreach ($k in $cur.Keys) {
    $h = Get-TextHash $cur[$k]
    if (-not $baseHash.ContainsKey($k)) { $added += $k; $take += $k }
    elseif ($baseHash[$k] -ne $h) { $take += $k }
}
$removed = @($baseHash.Keys | Where-Object { -not $cur.Contains($_) })

if ($take.Count -eq 0) {
    Write-Host ''
    Write-Host 'После последней загрузки объекты не менялись - слать нечего.' -ForegroundColor Yellow
    exit 0
}
foreach ($k in $take) {
    $head = ($cur[$k] -split "`r`n", 2)[0]
    Step ('  ' + $head.Substring(7) + $(if ($added -contains $k) { '   (новый)' } else { '' }))
}
if ($removed.Count -gt 0) {
    # Частичный импорт умеет заменять и добавлять, но не удалять. Молчать об этом нельзя.
    foreach ($k in $removed) { Write-Host ("  УДАЛЁН, импортом не снимается: $k") -ForegroundColor Yellow }
}

$stamp = (Get-Date -Format 'yyyyMMdd')
$body = ($take | ForEach-Object { $cur[$_] }) -join ''

# --- голые объекты -----------------------------------------------------------
if ($ObjectsOnly) {
    # Отправляемый файл лежит ОДИН и в собственном каталоге. Пока соседями по out\ были
    # вчерашняя дельта и utf8-близнец "для чтения", выбор делался глазами по имени - и
    # однажды уехал не тот файл. Каталоги чистятся при КАЖДОЙ сборке, поэтому правило
    # звучит без оговорок: отправлять всё, что лежит в out\send, а там всегда один файл.
    $send = Join-Path $OutDir 'send'
    $read = Join-Path $OutDir 'read'
    foreach ($d in @($send, $read)) {
        if (Test-Path $d) { Get-ChildItem -LiteralPath $d -Force | Remove-Item -Force -Recurse }
        else { New-Item -ItemType Directory -Path $d | Out-Null }
    }
    # Прежние плоские файлы из корня out\ сносим тем же заходом: оставшись лежать рядом,
    # они снова стали бы кандидатами на отправку.
    @(Get-ChildItem -LiteralPath $OutDir -Filter 'LineProfiler-changed-*' -File) |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }

    # Сборка идёт в каталоге для чтения: convert-for-cside кладёт cp866 рядом с исходным
    # файлом, и уже готовый переносится в send. Так в send не появляется ничего лишнего
    # даже на миг.
    $txt = Join-Path $read ("LineProfiler-changed-$stamp.txt")
    [IO.File]::WriteAllText($txt, $body, [Text.UTF8Encoding]::new($false))
    & (Join-Path $PSScriptRoot 'convert-for-cside.ps1') -TaskFile $txt | Out-Null
    $made = Join-Path $read ("LineProfiler-changed-$stamp.cp866.txt")
    if (-not (Test-Path $made)) { throw 'не собрался cp866' }
    $cp = Join-Path $send ("LineProfiler-changed-$stamp.cp866.txt")
    Move-Item -LiteralPath $made -Destination $cp -Force

    # Сторож на само правило: если в send оказалось не одно, отправлять снова наугад.
    $inSend = @(Get-ChildItem -LiteralPath $send -Force)
    if ($inSend.Count -ne 1) { throw ('в out\send файлов ' + $inSend.Count + ', а должен быть ровно один') }

    Write-Host ''
    Write-Host 'ИТОГ' -ForegroundColor Green
    Step ('ОТПРАВЛЯТЬ: out\send\' + (Split-Path $cp -Leaf) + "   {0:N0} байт" -f (Get-Item $cp).Length)
    Step ('в out\send это единственный файл - брать можно не глядя на имя')
    Step ('для чтения: out\read\' + (Split-Path $txt -Leaf) + "   {0:N0} байт, UTF-8" -f (Get-Item $txt).Length)
    exit 0
}

# --- пакет -------------------------------------------------------------------
$stage = Join-Path $env:TEMP ('lp-delta-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'objects') | Out-Null
$deltaTxt = Join-Path $stage 'objects\LineProfiler.txt'
[IO.File]::WriteAllText($deltaTxt, $body, [Text.UTF8Encoding]::new($false))
& (Join-Path $PSScriptRoot 'convert-for-cside.ps1') -TaskFile $deltaTxt | Out-Null
if (-not (Test-Path (Join-Path $stage 'objects\LineProfiler.cp866.txt'))) { throw 'не собрался cp866' }

# Прочие файлы кладём, только когда есть с чем сверить - то есть при -Previous.
$plain = @()
$dllSent = $false
if ($Previous) {
    $onsite = Join-Path $repo 'onsite'
    foreach ($f in (Get-ChildItem $onsite -File | Where-Object { $_.Extension -in @('.ps1', '.md') })) {
        if (($prevHash[$f.Name]) -ne (Get-FileHash $f.FullName -Algorithm SHA256).Hash) {
            Copy-Item $f.FullName $stage
            $plain += $f.Name
        }
    }
    # Сборка приёмника - самый тяжёлый файл. Сверяем с тем, что УЕХАЛО, а не с локальной
    # пересборкой: csc недетерминирован, и две сборки одного исходника всегда разные.
    $dll = Join-Path $repo 'dist\AlLineProfiler.dll'
    if ((Test-Path $dll) -and ($prevHash['receiver\AlLineProfiler.dll'] -ne (Get-FileHash $dll -Algorithm SHA256).Hash)) {
        New-Item -ItemType Directory -Path (Join-Path $stage 'receiver') | Out-Null
        Copy-Item $dll (Join-Path $stage 'receiver')
        Copy-Item (Join-Path $repo 'src\AlLineProfiler.cs') (Join-Path $stage 'receiver')
        $dllSent = $true
    }
}

$commit = (& git -C $repo rev-parse --short HEAD 2>$null)
$now = (Get-Date -Format 'yyyy-MM-dd HH:mm')
$apply = @(
    'КАК ПРИМЕНИТЬ',
    '',
    'Это ИЗМЕНЕНИЯ. Они дополняют установленную поставку, а не заменяют её.',
    '',
    ('1. C/SIDE -> File -> Import -> objects\LineProfiler.cp866.txt (объектов: ' + $take.Count + ').'),
    '   C/SIDE заменит ровно эти объекты, остальные не тронет.',
    '2. Скомпилировать импортированные объекты.',
    '3. Закрыть и открыть заново клиент. Codeunit 110200 объявлен SingleInstance,',
    '   и открытая сессия держит прежний экземпляр до переподключения.',
    ''
)
if ($dllSent) {
    $apply += @('СБОРКА ПРИЁМНИКА ИЗМЕНИЛАСЬ - выложить шагом 03-Build-Receiver.ps1.', '')
} else {
    $apply += @('Сборка приёмника не менялась: Add-ins не трогается, служба не перезапускается.', '')
}
$apply += @('ПРОВЕРКА', '  Настройка -> галка "Обкатка" -> Codeunit 110200 -> Run.',
            '  Первая строка вердикта: passed N of M.', '')
[IO.File]::WriteAllLines((Join-Path $stage 'APPLY.txt'), $apply, [Text.UTF8Encoding]::new($true))

$ch = @('Что изменилось', '', ("Собран $now, коммит $commit"), ("Сравнение с: $baseName"), '', 'ОБЪЕКТЫ:')
foreach ($k in $take) {
    $head = ($cur[$k] -split "`r`n", 2)[0]
    $ch += ('  ' + $head.Substring(7) + $(if ($added -contains $k) { '   (новый)' } else { '' }))
}
$ch += ''
$ch += ('Не изменились и не вошли: ' + ($cur.Count - $take.Count) + ' объектов.')
if ($removed.Count -gt 0) {
    $ch += ''
    $ch += 'УДАЛЕНЫ ИЗ ПОСТАВКИ (импортом НЕ снимаются, снять вручную в C/SIDE):'
    foreach ($k in $removed) { $ch += ('  ' + $k) }
}
if ($plain.Count -gt 0) {
    $ch += ''
    $ch += 'ПРОЧИЕ ФАЙЛЫ:'
    foreach ($n in $plain) { $ch += ('  ' + $n) }
}
[IO.File]::WriteAllLines((Join-Path $stage 'CHANGES.txt'), $ch, [Text.UTF8Encoding]::new($true))

$lines = @('Пакет изменений: построчный профайлер C/AL', "Собран: $now", "Исходник: коммит $commit",
           "Сравнение с: $baseName", '', 'Файл                                     Размер  SHA256')
foreach ($f in (Get-ChildItem $stage -Recurse -File | Sort-Object FullName)) {
    $rel = $f.FullName.Substring($stage.Length + 1)
    $lines += ('{0,-40} {1,7}  {2}' -f $rel, $f.Length, (Get-FileHash $f.FullName -Algorithm SHA256).Hash.Substring(0, 16))
}
[IO.File]::WriteAllLines((Join-Path $stage 'MANIFEST.txt'), $lines, [Text.UTF8Encoding]::new($true))

# Каталоги те же, что и у голых объектов: отправляемое - в send, всё прочее - в read.
$send = Join-Path $OutDir 'send'
$read = Join-Path $OutDir 'read'
foreach ($d in @($send, $read)) {
    if (Test-Path $d) { Get-ChildItem -LiteralPath $d -Force | Remove-Item -Force -Recurse }
    else { New-Item -ItemType Directory -Path $d | Out-Null }
}
@(Get-ChildItem -LiteralPath $OutDir -Filter 'LineProfiler-delta-*' -File) |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }

$zip = Join-Path $read ("LineProfiler-delta-$stamp.zip")
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

# Тело письма собирается ТЕМ ЖЕ заходом. Почта контура вложений не принимает, поэтому
# уезжает не архив, а его текст; сам архив остаётся в read - с ним сверяют собранное на
# той стороне. Вторым шагом руками это забывалось, и в out\ лежал архив, который
# отправить нельзя, рядом с текстом, который отправить можно.
& (Join-Path $PSScriptRoot 'ConvertTo-MailBase64.ps1') -Path $zip -OutDir $send -PartChars $PartChars
$letters = @(Get-ChildItem -LiteralPath $send -Force | Sort-Object Name)
if ($letters.Count -lt 1) { throw 'тело письма не собралось' }

Write-Host ''
Write-Host 'ИТОГ' -ForegroundColor Green
Step ('ОТПРАВЛЯТЬ: всё из out\send, писем ' + $letters.Count)
foreach ($l in $letters) { Step ('  ' + $l.Name + "   {0:N0} байт" -f $l.Length) }
Step ('архив для сверки: out\read\' + (Split-Path $zip -Leaf) + "   {0:N0} байт" -f (Get-Item $zip).Length)
