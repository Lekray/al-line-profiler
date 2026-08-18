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

.EXAMPLE
    pwsh scripts/New-OnsitePackage.ps1
#>
[CmdletBinding()]
param([string]$OutDir)

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
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

$size = [math]::Round((Get-Item $zip).Length / 1KB)
Write-Host ''
Write-Host "ИТОГ: $zip  ($size КБ)" -ForegroundColor Green
