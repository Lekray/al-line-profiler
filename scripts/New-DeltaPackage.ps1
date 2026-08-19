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
    Отдать ГОЛЫЕ объекты: <имя>.cp866.txt для импорта и <имя>.txt для чтения. Без zip,
    без README и без манифеста - когда нужен просто файл, который импортируют.

.PARAMETER Previous
    Пакет, с которым сравнивать, вместо delivered.txt.

.PARAMETER Record
    Записать delivered.txt по текущему тексту: считать, что эти объекты загружены.

.PARAMETER OutDir
    Куда положить. По умолчанию out\ в корне репозитория.

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
    [switch] $Record
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
    $txt = Join-Path $OutDir ("LineProfiler-changed-$stamp.txt")
    [IO.File]::WriteAllText($txt, $body, [Text.UTF8Encoding]::new($false))
    & (Join-Path $PSScriptRoot 'convert-for-cside.ps1') -TaskFile $txt | Out-Null
    $cp = Join-Path $OutDir ("LineProfiler-changed-$stamp.cp866.txt")
    if (-not (Test-Path $cp)) { throw 'не собрался cp866' }
    Write-Host ''
    Write-Host 'ИТОГ' -ForegroundColor Green
    Step ((Split-Path $cp -Leaf)  + "   {0:N0} байт - ИМПОРТИРОВАТЬ ЭТОТ" -f (Get-Item $cp).Length)
    Step ((Split-Path $txt -Leaf) + "   {0:N0} байт - тот же текст в UTF-8, для чтения" -f (Get-Item $txt).Length)
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

$zip = Join-Path $OutDir ("LineProfiler-delta-$stamp.zip")
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

Write-Host ''
Write-Host ("ИТОГ: $zip  ({0:N0} КБ)" -f [math]::Round((Get-Item $zip).Length / 1KB)) -ForegroundColor Green
