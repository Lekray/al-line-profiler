#requires -Version 7
<#
.SYNOPSIS
    Пакет ТОЛЬКО изменений против пакета, который у получателя уже есть.

.DESCRIPTION
    Полный пакет везёт все 16 объектов - это 444 КБ одного лишь cp866, а телом письма
    выходит около 400 КБ base64. Когда на объекте уже стоит прошлая поставка, столько
    везти незачем: правится обычно один-два объекта.

    Поэтому текст объектов режется на объекты и сравнивается с тем, что лежал в прошлом
    пакете. В новый файл попадают только разошедшиеся - C/SIDE принимает частичный
    текстовый экспорт и заменяет ровно те объекты, что в нём есть.

    Остальные файлы (шаги установки, README, сборка приёмника) сверяются целиком по
    SHA-256 и кладутся, только если изменились.

    Пакет НЕ самодостаточен: он дополняет прошлый, а не заменяет его. Тем, у кого
    прошлого нет, нужен полный - New-OnsitePackage.ps1.

.PARAMETER Previous
    Пакет, который у получателя уже есть. Именно с ним и идёт сравнение.

.PARAMETER OutDir
    Куда положить. По умолчанию out\ в корне репозитория.

.EXAMPLE
    pwsh scripts/New-DeltaPackage.ps1 -Previous out/LineProfiler-onsite-20260818.zip
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $Previous,
    [string] $OutDir
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
if (-not (Test-Path -LiteralPath $Previous)) { throw "не найден прошлый пакет: $Previous" }

function Step([string]$m) { Write-Host "  $m" }
Write-Host 'Пакет изменений' -ForegroundColor Cyan

# --- что лежало в прошлом пакете --------------------------------------------
$prevZip  = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Previous).Path)
$prevText = ''
$prevHash = @{}
try {
    foreach ($e in $prevZip.Entries) {
        if ($e.FullName -like '*/') { continue }
        $ms = New-Object IO.MemoryStream
        $s = $e.Open(); $s.CopyTo($ms); $s.Close()
        $bytes = $ms.ToArray(); $ms.Dispose()
        $key = $e.FullName -replace '/', '\'
        $prevHash[$key] = [BitConverter]::ToString(
            [Security.Cryptography.SHA256]::Create().ComputeHash($bytes)).Replace('-', '')
        if ($key -eq 'objects\LineProfiler.txt') {
            # Текст объектов читаем БЕЗ BOM и как есть: сравнивать будем посимвольно.
            $prevText = [Text.Encoding]::UTF8.GetString($bytes)
        }
    }
} finally { $prevZip.Dispose() }
if (-not $prevText) { throw "в прошлом пакете нет objects\LineProfiler.txt - разницу по объектам не с чем считать" }
Step ("прошлый пакет: " + (Split-Path $Previous -Leaf) + ", файлов " + $prevHash.Count)

# --- режем текст на объекты --------------------------------------------------
# Объект начинается со строки OBJECT в первой позиции и тянется до следующей такой.
# Иначе резать нечем: у C/SIDE нет ни разделителя, ни оглавления в текстовом экспорте.
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

$curFile = Join-Path $repo 'LineProfiler.txt'
$curText = [IO.File]::ReadAllText($curFile, [Text.UTF8Encoding]::new($false))
$cur  = Split-CSideObjects $curText
$prev = Split-CSideObjects $prevText
Step ("объектов: было $($prev.Count), стало $($cur.Count)")

$changed = @()
$added   = @()
foreach ($k in $cur.Keys) {
    if (-not $prev.Contains($k)) { $added += $k }
    elseif ($cur[$k] -ne $prev[$k]) { $changed += $k }
}
$removed = @($prev.Keys | Where-Object { -not $cur.Contains($_) })
$take = @($cur.Keys | Where-Object { ($changed -contains $_) -or ($added -contains $_) })

if ($take.Count -eq 0) {
    Write-Host ''
    Write-Host 'Объекты не менялись - слать нечего.' -ForegroundColor Yellow
} else {
    Step ("изменились: " + ($take -join ', '))
}

# --- стенд пакета ------------------------------------------------------------
$stage = Join-Path $env:TEMP ('lp-delta-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'objects') | Out-Null

if ($take.Count -gt 0) {
    $deltaTxt = Join-Path $stage 'objects\LineProfiler.txt'
    # Порядок и байты - как в исходном файле: части склеиваются встык, каждая уже
    # заканчивается переводом строки. Ни BOM, ни хвоста мы не добавляем: C/SIDE
    # спотыкается о лишние байты перед OBJECT.
    $body = ($take | ForEach-Object { $cur[$_] }) -join ''
    [IO.File]::WriteAllText($deltaTxt, $body, [Text.UTF8Encoding]::new($false))
    & (Join-Path $PSScriptRoot 'convert-for-cside.ps1') -TaskFile $deltaTxt | Out-Null
    $deltaCp = Join-Path $stage 'objects\LineProfiler.cp866.txt'
    if (-not (Test-Path $deltaCp)) { throw 'не собрался cp866 для пакета изменений' }
    Step ("объекты: {0} КБ текста, {1} КБ cp866" -f
          [math]::Round((Get-Item $deltaTxt).Length / 1KB), [math]::Round((Get-Item $deltaCp).Length / 1KB))
}

# --- прочие файлы: только те, что разошлись ----------------------------------
$onsite = Join-Path $repo 'onsite'
$plain = @()
foreach ($f in (Get-ChildItem $onsite -File | Where-Object { $_.Extension -in @('.ps1', '.md') })) {
    $h = (Get-FileHash $f.FullName -Algorithm SHA256).Hash
    if (($prevHash[$f.Name]) -ne $h) {
        Copy-Item $f.FullName $stage
        $plain += $f.Name
    }
}
# Сборка приёмника - самый тяжёлый файл пакета. Кладём, только если она правда другая:
# csc недетерминирован, поэтому сверяем с тем, что УЕХАЛО, а не с локальной пересборкой.
$dll = Join-Path $repo 'dist\AlLineProfiler.dll'
$dllSent = $false
if (Test-Path $dll) {
    $h = (Get-FileHash $dll -Algorithm SHA256).Hash
    if ($prevHash['receiver\AlLineProfiler.dll'] -ne $h) {
        New-Item -ItemType Directory -Path (Join-Path $stage 'receiver') | Out-Null
        Copy-Item $dll (Join-Path $stage 'receiver')
        Copy-Item (Join-Path $repo 'src\AlLineProfiler.cs') (Join-Path $stage 'receiver')
        $dllSent = $true
    }
}
Step ("прочих файлов: " + $plain.Count + $(if ($dllSent) { ', сборка приёмника ИЗМЕНИЛАСЬ' } else { ', сборка приёмника прежняя' }))

# --- как это применить -------------------------------------------------------
$commit = (& git -C $repo rev-parse --short HEAD 2>$null)
$stamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
$apply = @(
    'КАК ПРИМЕНИТЬ',
    '',
    'Это пакет ИЗМЕНЕНИЙ. Он дополняет уже установленную поставку, а не заменяет её:',
    'здесь лежат только те объекты и файлы, которые разошлись с прошлым пакетом.',
    ''
)
if ($take.Count -gt 0) {
    $apply += @(
        ('1. C/SIDE -> File -> Import -> objects\LineProfiler.cp866.txt (объектов: ' + $take.Count + ').'),
        '   C/SIDE заменит ровно эти объекты, остальные не тронет.',
        '2. Скомпилировать импортированные объекты.',
        '3. Закрыть и открыть заново клиент. Codeunit 110200 объявлен SingleInstance,',
        '   и открытая сессия продолжает держать прежний экземпляр до переподключения.',
        '   Если страница ведёт себя по-старому - перезапустить экземпляр службы.',
        ''
    )
}
if ($dllSent) {
    $apply += @(
        'СБОРКА ПРИЁМНИКА ИЗМЕНИЛАСЬ - её тоже надо выложить:',
        '   шаг 03-Build-Receiver.ps1 из полного пакета, либо вручную положить',
        '   receiver\AlLineProfiler.dll в Add-ins\LineProfiler и перезапустить службу.',
        ''
    )
} else {
    $apply += @(
        'Сборка приёмника НЕ менялась: ни выкладывать её, ни перезапускать службу ради',
        'неё не нужно. Add-ins не трогается вовсе.',
        ''
    )
}
$apply += @(
    'ЧТО ПРОВЕРИТЬ ПОСЛЕ',
    '  Настройка профайлера -> галка "Обкатка" -> Object Designer -> Codeunit 110200 ->',
    '  Run. Первая строка вердикта - "passed N of M". Эталон: 34 из 34.',
    ''
)
[IO.File]::WriteAllLines((Join-Path $stage 'APPLY.txt'), $apply, [Text.UTF8Encoding]::new($true))

# --- что именно изменилось ---------------------------------------------------
$ch = @(
    'Что изменилось',
    '',
    ('Пакет изменений: собран ' + $stamp + ', коммит ' + $commit),
    ('Дополняет:       ' + (Split-Path $Previous -Leaf)),
    ''
)
if ($take.Count -gt 0) {
    $ch += 'ОБЪЕКТЫ В ПАКЕТЕ:'
    foreach ($k in $take) {
        $head = ($cur[$k] -split "`r`n", 2)[0]
        $ch += ('  ' + $head.Substring(7) + $(if ($added -contains $k) { '   (новый)' } else { '' }))
    }
    $ch += ''
    $ch += ('Не изменились и в пакет не вошли: ' + ($cur.Count - $take.Count) + ' объектов.')
    $ch += ''
}
if ($removed.Count -gt 0) {
    # Частичный импорт умеет заменять и добавлять, но не удалять. Молчать об этом нельзя.
    $ch += 'УДАЛЕНЫ ИЗ ПОСТАВКИ (импортом НЕ снимаются, снять вручную в C/SIDE):'
    foreach ($k in $removed) { $ch += ('  ' + $k) }
    $ch += ''
}
if ($plain.Count -gt 0) {
    $ch += 'ПРОЧИЕ ФАЙЛЫ:'
    foreach ($n in $plain) { $ch += ('  ' + $n) }
    $ch += ''
}
$ch += 'Сверялось содержимое: объекты - посимвольно, прочие файлы - по SHA-256.'
[IO.File]::WriteAllLines((Join-Path $stage 'CHANGES.txt'), $ch, [Text.UTF8Encoding]::new($true))

# --- манифест ----------------------------------------------------------------
$lines = @(
    'Пакет изменений: построчный профайлер C/AL',
    ("Собран: $stamp"),
    ("Исходник: коммит $commit"),
    ("Дополняет: " + (Split-Path $Previous -Leaf)),
    '',
    'Файл                                     Размер  SHA256'
)
foreach ($f in (Get-ChildItem $stage -Recurse -File | Sort-Object FullName)) {
    $rel = $f.FullName.Substring($stage.Length + 1)
    $h = (Get-FileHash $f.FullName -Algorithm SHA256).Hash.Substring(0, 16)
    $lines += ('{0,-40} {1,7}  {2}' -f $rel, $f.Length, $h)
}
[IO.File]::WriteAllLines((Join-Path $stage 'MANIFEST.txt'), $lines, [Text.UTF8Encoding]::new($true))

$zip = Join-Path $OutDir ('LineProfiler-delta-' + (Get-Date -Format 'yyyyMMdd') + '.zip')
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

$size = [math]::Round((Get-Item $zip).Length / 1KB)
Write-Host ''
Write-Host "ИТОГ: $zip  ($size КБ)" -ForegroundColor Green
