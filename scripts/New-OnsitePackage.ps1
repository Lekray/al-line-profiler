#requires -Version 7
<#
.SYNOPSIS
    Собирает ZIP для отправки на сервер на объекте: объекты, приёмник, шаги диагностики.

.DESCRIPTION
    Заливать в изолированный контур получается редко, поэтому пакет должен быть полным
    с первого раза. Скрипт берёт ТЕКУЩИЕ исходники, пересобирает cp866 и приёмник и
    складывает всё вместе с манифестом.

    Приёмник кладётся как ЗАПАСНОЙ вариант: на месте его пересобирают шагом 3, потому
    что ссылка идёт на ту TraceEvent.dll, что лежит там, а её версия своя.

    Рядом с манифестом кладётся CHANGES.txt - что изменилось относительно ПРОШЛОГО
    пакета. Согласующему иначе приходится сличать манифесты глазами, а пакет уходит
    целиком и всегда: частичная доставка в изолированный контур - это установка,
    собранная из двух источников, и разбираться, что там от какой версии, будет уже
    человек на объекте.

.PARAMETER Previous
    Прошлый пакет, с которым сравнивать. По умолчанию берётся самый свежий
    LineProfiler-onsite-*.zip в выходном каталоге, кроме собираемого сейчас.

.PARAMETER NoChanges
    Не сравнивать с прошлым пакетом и не класть CHANGES.txt.

.EXAMPLE
    pwsh scripts/New-OnsitePackage.ps1
#>
[CmdletBinding()]
param([string]$OutDir, [string]$Previous, [switch]$NoChanges)

$ErrorActionPreference = 'Stop'

$taskDir   = Split-Path $PSScriptRoot -Parent
$onsiteDir = Join-Path $taskDir 'onsite'
if (-not $OutDir) { $OutDir = Join-Path $taskDir 'out' }

function Step([string]$m) { Write-Host "  $m" }
Write-Host 'Пакет для сервера на объекте' -ForegroundColor Cyan

# 1. cp866 пересобираем ВСЕГДА: файл задачи мог поменяться после прошлой сборки, а
#    импорт в C/SIDE другой кодировки не принимает - кириллица приезжает мозаикой.
$conv = Join-Path $PSScriptRoot 'convert-for-cside.ps1'
& $conv -TaskFile (Join-Path $taskDir 'LineProfiler.txt') | Out-Null
$cp866 = Join-Path $taskDir 'LineProfiler.cp866.txt'
if (-not (Test-Path $cp866)) { throw 'не собрался cp866' }
Step 'cp866 пересобран'

# В пакет идёт то, что ОПУБЛИКОВАНО в dist: именно по нему согласована установка и
# посчитаны контрольные суммы. Свежая локальная сборка в bin - повод обновить dist,
# а не тихо подменить содержимое пакета.
$dll = Join-Path $taskDir 'dist\AlLineProfiler.dll'
if (-not (Test-Path $dll)) { throw "нет опубликованной сборки: $dll" }
$bin = Join-Path $taskDir 'bin\AlLineProfiler.dll'
if ((Test-Path $bin) -and ((Get-FileHash $bin).Hash -ne (Get-FileHash $dll).Hash)) {
    throw "bin\AlLineProfiler.dll отличается от dist\AlLineProfiler.dll. В пакет идёт dist. Обновите dist и пересчитайте dist\SHA256SUMS.txt либо удалите bin."
}
$variants = Join-Path $taskDir 'dist\variants'

$stage = Join-Path $env:TEMP ('lineprofiler-package-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'objects')  | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'receiver') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'receiver\variants') | Out-Null

Copy-Item (Join-Path $onsiteDir '*.ps1') $stage
Copy-Item (Join-Path $onsiteDir 'README.md') $stage
Copy-Item $cp866 (Join-Path $stage 'objects')
Copy-Item (Join-Path $taskDir 'LineProfiler.txt') (Join-Path $stage 'objects')
Copy-Item (Join-Path $taskDir 'src\AlLineProfiler.cs') (Join-Path $stage 'receiver')
Copy-Item $dll (Join-Path $stage 'receiver')
# Варианты той же нашей сборки под другие версии TraceEvent - на случай сервера
# без компилятора. Саму TraceEvent не везём: она уже на сервере, и её лицензия
# распространение не разрешает (см. THIRD-PARTY-NOTICES.md).
if (Test-Path $variants) { Copy-Item (Join-Path $variants '*.dll') (Join-Path $stage 'receiver\variants') }
Step 'состав собран'

# 2. Манифест: в контуре нет ни git, ни сети, и по нему на месте видно, что именно
#    приехало и не побилось ли по дороге.
$commit = (& git -C $taskDir rev-parse --short HEAD 2>$null)
$stamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
$lines = @(
    'Пакет: построчный профайлер C/AL',
    "Собран: $stamp",
    "Исходник: коммит $commit",
    '',
    'Файл                                     Размер  SHA256'
)
foreach ($f in (Get-ChildItem $stage -Recurse -File | Sort-Object FullName)) {
    $rel = $f.FullName.Substring($stage.Length + 1)
    $h = (Get-FileHash $f.FullName -Algorithm SHA256).Hash.Substring(0, 16)
    $lines += ('{0,-40} {1,7}  {2}' -f $rel, $f.Length, $h)
}
$manifest = Join-Path $stage 'MANIFEST.txt'
[IO.File]::WriteAllLines($manifest, $lines, [Text.UTF8Encoding]::new($true))
Step 'манифест записан'

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$zip = Join-Path $OutDir ('LineProfiler-onsite-' + (Get-Date -Format 'yyyyMMdd') + '.zip')

# 3. Что изменилось с прошлого пакета. Сравниваются СОДЕРЖИМЫЕ файлов по хэшу, а не
#    даты: дата меняется от пересборки cp866 и ничего не значит. MANIFEST.txt и сам
#    CHANGES.txt из сравнения исключены - они пересобираются каждый раз и в списке
#    изменений стояли бы всегда, обесценивая его.
if (-not $NoChanges) {
    if (-not $Previous) {
        $cand = @(Get-ChildItem (Join-Path $OutDir 'LineProfiler-onsite-*.zip') -ErrorAction SilentlyContinue |
                  Where-Object { $_.FullName -ne $zip } | Sort-Object Name -Descending)
        if ($cand.Count -gt 0) { $Previous = $cand[0].FullName }
    }
    if (-not $Previous -or -not (Test-Path -LiteralPath $Previous)) {
        Step 'прошлого пакета рядом нет - CHANGES.txt не кладётся'
    }
    else {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $prevMap = @{}
        $za = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $Previous).Path)
        try {
            foreach ($e in $za.Entries) {
                if (-not $e.Name) { continue }                       # каталог
                $s = $e.Open()
                try {
                    $ms = New-Object System.IO.MemoryStream
                    $s.CopyTo($ms)
                    $sha = [System.Security.Cryptography.SHA256]::Create()
                    try { $h = ([BitConverter]::ToString($sha.ComputeHash($ms.ToArray()))).Replace('-','') }
                    finally { $sha.Dispose() }
                    $prevMap[$e.FullName.Replace('/', '\')] = @{ Size = $e.Length; Hash = $h }
                }
                finally { $s.Dispose() }
            }
        }
        finally { $za.Dispose() }

        $skip = @('MANIFEST.txt', 'CHANGES.txt')
        $now  = @{}
        foreach ($f in (Get-ChildItem $stage -Recurse -File)) {
            $rel = $f.FullName.Substring($stage.Length + 1)
            if ($skip -contains $rel) { continue }
            $now[$rel] = @{ Size = $f.Length; Hash = (Get-FileHash $f.FullName -Algorithm SHA256).Hash }
        }
        foreach ($k in @($prevMap.Keys)) { if ($skip -contains $k) { $prevMap.Remove($k) } }

        $added = @(); $changed = @(); $removed = @(); $same = 0
        foreach ($k in ($now.Keys | Sort-Object)) {
            if (-not $prevMap.ContainsKey($k)) { $added += $k }
            elseif ($prevMap[$k].Hash -ne $now[$k].Hash) { $changed += $k }
            else { $same++ }
        }
        foreach ($k in ($prevMap.Keys | Sort-Object)) { if (-not $now.ContainsKey($k)) { $removed += $k } }

        $ch = @(
            'Что изменилось в пакете',
            '',
            ('Этот пакет:   {0}, собран {1}, коммит {2}' -f (Split-Path $zip -Leaf), $stamp, $commit),
            ('Сравнение с:  {0}' -f (Split-Path $Previous -Leaf)),
            '',
            'Сверялось содержимое файлов по SHA-256, а не даты: дата меняется от',
            'пересборки и сама по себе ничего не значит. MANIFEST.txt и этот файл',
            'в сравнение не входят - они пересобираются каждый раз.',
            ''
        )
        if ($added.Count -eq 0 -and $changed.Count -eq 0 -and $removed.Count -eq 0) {
            $ch += 'Содержимое пакета не изменилось ни одним файлом.'
        }
        else {
            if ($changed.Count -gt 0) {
                $ch += ('ИЗМЕНЕНЫ ({0}):' -f $changed.Count)
                foreach ($k in $changed) {
                    $ch += ('  {0,-44} {1,8} байт  {2}' -f $k, $now[$k].Size, $now[$k].Hash.Substring(0, 16))
                    $ch += ('  {0,-44} {1,8} байт  {2}   было' -f '', $prevMap[$k].Size, $prevMap[$k].Hash.Substring(0, 16))
                }
                $ch += ''
            }
            if ($added.Count -gt 0) {
                $ch += ('ДОБАВЛЕНЫ ({0}):' -f $added.Count)
                foreach ($k in $added) {
                    $ch += ('  {0,-44} {1,8} байт  {2}' -f $k, $now[$k].Size, $now[$k].Hash.Substring(0, 16))
                }
                $ch += ''
            }
            if ($removed.Count -gt 0) {
                $ch += ('УДАЛЕНЫ ({0}):' -f $removed.Count)
                foreach ($k in $removed) { $ch += ('  {0}' -f $k) }
                $ch += ''
            }
        }
        $ch += ('Без изменений: {0} файлов.' -f $same)
        [IO.File]::WriteAllLines((Join-Path $stage 'CHANGES.txt'), $ch, [Text.UTF8Encoding]::new($true))
        Step ('CHANGES.txt: изменено {0}, добавлено {1}, удалено {2}, без изменений {3}' -f
              $changed.Count, $added.Count, $removed.Count, $same)
    }
}
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

$size = [math]::Round((Get-Item $zip).Length / 1KB)
Write-Host ''
Write-Host "ИТОГ: $zip  ($size КБ)" -ForegroundColor Green
