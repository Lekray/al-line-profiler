#requires -Version 7
<#
.SYNOPSIS
    Собирает пакет В ВИДЕ ТЕКСТА - для отправки телом письма.

.DESCRIPTION
    Почтовый фильтр контура удаляет вложения, поэтому пакет едет текстом: base64
    строками по 64 знака, которые переживают любой перенос строк в почтовом клиенте.

    Поток РОВНО ОДИН - архив, внутри которого лежит и сборка приёмника. Вторым потоком
    её везти нельзя: два документа base64 подряд - это уже не один документ base64.
    Исходник в UTF-8 не кладётся: для импорта нужен cp866, а править объекты на месте
    удобнее уже в C/SIDE.

    На выходе - и один файл целиком, и он же нарезанный на части: если тело письма
    целиком не пройдёт, части уходят по отдельности и собираются в любом порядке.

.PARAMETER PartChars
    Знаков полезной нагрузки в одной части. По умолчанию 44800 (кратно 64) - около 46 КБ тела.

.EXAMPLE
    pwsh scripts/New-MailPackage.ps1
#>
[CmdletBinding()]
param([int]$PartChars = 44800, [string]$OutDir, [switch]$UnpackerOnly)

$ErrorActionPreference = 'Stop'

$taskDir   = Split-Path $PSScriptRoot -Parent
$onsiteDir = Join-Path $taskDir 'onsite'
if (-not $OutDir) { $OutDir = Join-Path $taskDir 'out' }

function Step([string]$m) { Write-Host "  $m" }
Write-Host 'Почтовый пакет' -ForegroundColor Cyan

# cp866 пересобираем всегда: файл задачи мог поменяться, а C/SIDE другой кодировки
# не принимает - кириллица приезжает мозаикой.
& (Join-Path $PSScriptRoot 'convert-for-cside.ps1') -TaskFile (Join-Path $taskDir 'LineProfiler.txt') | Out-Null
$cp866 = Join-Path $taskDir 'LineProfiler.cp866.txt'
if (-not (Test-Path $cp866)) { throw 'не собрался cp866' }
Step 'cp866 пересобран'

$stage = Join-Path $env:TEMP ('lineprofiler-mail-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'objects')  | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stage 'receiver') | Out-Null
Copy-Item (Join-Path $onsiteDir '*.ps1') $stage
Copy-Item (Join-Path $onsiteDir 'README.md') $stage
Copy-Item $cp866 (Join-Path $stage 'objects')
Copy-Item (Join-Path $taskDir 'src\AlLineProfiler.cs') (Join-Path $stage 'receiver')
# Сборка едет ВНУТРИ архива, рядом с исходником. Она нужна на РАБОЧЕЙ СТАНЦИИ: положить
# её на сервер сможет либо установщик из C/AL (если у учётки службы есть права на
# Add-ins), либо человек с доступом к серверу - но в обоих случаях файл сперва должен
# доехать сюда.
#
# Берём ОПУБЛИКОВАННУЮ сборку из dist, а не свежую из bin: именно по dist согласована
# установка и посчитаны контрольные суммы. Так же поступает и New-OnsitePackage.ps1.
$dll = Join-Path $taskDir 'dist\AlLineProfiler.dll'
if (-not (Test-Path $dll)) { throw "нет опубликованной сборки: $dll" }
$bin = Join-Path $taskDir 'bin\AlLineProfiler.dll'
if ((Test-Path $bin) -and ((Get-FileHash $bin).Hash -ne (Get-FileHash $dll).Hash)) {
    throw "bin\AlLineProfiler.dll отличается от dist\AlLineProfiler.dll. В пакет идёт dist. Обновите dist и пересчитайте dist\SHA256SUMS.txt либо удалите bin."
}
Copy-Item $dll (Join-Path $stage 'receiver')

$zip = Join-Path $env:TEMP 'lineprofiler-mail.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

$bytes = [IO.File]::ReadAllBytes($zip)
$hash  = (Get-FileHash $zip -Algorithm SHA256).Hash
# Поток один, и это не вкусовщина. Второй поток под своей меткой тут был и оказался
# ловушкой сразу с двух сторон: сборщик из шапки письма забирал только первую метку и
# сборку терял вовсе, а распаковщик на C/AL складывал ОБЕ метки в один буфер - выходил
# либо отказ на дополнении base64 (два размера архива из трёх), либо архив с приклеенным
# к нему хвостом. Проверено обратной сборкой на настоящих частях.
# Строки по 64 знака плюс метка - итого 65: заведомо короче даже 72, на которых
# переносит формат format=flowed. Перенос строки посреди нагрузки терял бы её молча:
# продолжение уехало бы на строку без метки, а такую строку отбор выбрасывает.
# Отбор ПО МЕТКЕ, а не по длине: у последней строки каждой части длина своя, и отбор
# по длине выбрасывал её - проверено обратной сборкой, сумма не сошлась.
$flat = [Convert]::ToBase64String($bytes)
Step ("архив $([math]::Round($bytes.Length/1KB)) КБ (со сборкой внутри) -> текст $([math]::Round($flat.Length/1KB)) КБ")

$commit = (& git -C $taskDir rev-parse --short HEAD 2>$null)
$stamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm')

function New-Header {
    param([int]$No, [int]$Of)
    $h = @()
    $h += "=== ПОСТРОЧНЫЙ ПРОФАЙЛЕР C/AL - ПОЧТОВЫЙ ПАКЕТ - ЧАСТЬ $No ИЗ $Of ==="
    $h += ''
    $h += "Собран $stamp, исходник - коммит $commit."
    $h += 'Это ТЕКСТ, а не вложение: почтовый фильтр вложения удаляет. Двоичного тут нет.'
    $h += ''
    if ($No -eq 1) {
        $h += 'ЧТО СДЕЛАТЬ'
        $h += '  1. Сохранить тело КАЖДОГО письма как part01.txt ... part{0:d2}.txt в один каталог.' -f $Of
        $h += '  2. Из этого каталога выполнить в PowerShell:'
        $h += ''
        $h += '     $t = Get-ChildItem part*.txt | Sort-Object Name | Get-Content'
        $h += '     $b = -join (($t -match ''^\|'') -replace ''^\|'','''')'
        $h += '     [IO.File]::WriteAllBytes("$PWD\pkg.zip", [Convert]::FromBase64String($b))'
        $h += '     (Get-FileHash .\pkg.zip -Algorithm SHA256).Hash'
        $h += '     Expand-Archive .\pkg.zip -DestinationPath .\LineProfiler'
        $h += ''
        $h += "  3. Сверить сумму - должна быть: $hash"
        $h += '     Не сошлась - какая-то часть пришла обрезанной, распаковку не продолжать.'
        $h += '  4. Дальше по README.md из архива: шаги 01..06 по порядку.'
        $h += ''
        $h += 'ВСЁ ЭТО ДЕЛАЕТСЯ НА ВАШЕЙ РАБОЧЕЙ СТАНЦИИ, НЕ НА СЕРВЕРЕ.'
        $h += '  Объекты C/SIDE импортирует из файла на ВАШЕЙ машине, поэтому доступ к диску'
        $h += '  сервера для них не нужен вовсе. На сервер обязана попасть ровно одна вещь -'
        $h += '  сборка приёмника receiver\AlLineProfiler.dll в каталог службы Add-ins.'
        $h += '  Она уже внутри архива - отдельно её раскодировать и собирать не нужно.'
        $h += ''
        $h += 'ЕСЛИ POWERSHELL НЕТ И НА РАБОЧЕЙ СТАНЦИИ'
        $h += '  Разбор base64 умеет штатный certutil, он есть в любой Windows:'
        $h += ('     copy /b ' + ((1..$Of | ForEach-Object { 'part{0:d2}.txt' -f $_ }) -join '+') + ' all.txt')
        $h += '     findstr /b "|" all.txt > marked.txt'
        $h += '     (убрать первый знак каждой строки любым редактором - это и есть base64)'
        $h += '     certutil -decode payload.b64 pkg.zip'
        $h += '  Дальше архив распаковывается Проводником.'
        $h += ''
        $h += 'Полезная нагрузка помечена вертикальной чертой в начале строки - по ней она и'
        $h += 'отбирается, а вся эта шапка отсеивается сама. Строки короткие (65 знаков), чтобы'
        $h += 'никакой почтовый клиент не переносил их и не рвал нагрузку пополам.'
        $h += ''
    }
    $h += '<<< НАЧАЛО ЧАСТИ ' + $No + ' >>>'
    return $h
}

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
if (-not $UnpackerOnly) {
    # Через -Filter нельзя: он уходит в файловый API Windows, а тот знает только * и ?
    # - набор [0-9] там не работает вовсе, и части прошлой рассылки оставались лежать.
    # Лишняя часть, подобранная сборщиком вместе с новыми, ломает сборку молча.
    Get-ChildItem (Join-Path $OutDir 'LineProfiler-mail-[0-9][0-9].txt') -ErrorAction SilentlyContinue | Remove-Item -Force
}

# Сперва собираем ВСЕ строки нагрузки, потом режем по строкам: так ни одна строка не
# разрывается на границе части, а метка позволяет собирать части в любом порядке.
$payload = New-Object System.Collections.Generic.List[string]
for ($j = 0; $j -lt $flat.Length; $j += 64) {
    $payload.Add('|' + $flat.Substring($j, [math]::Min(64, $flat.Length - $j)))
}

$perPart = [math]::Floor($PartChars / 64)
$total = [math]::Ceiling($payload.Count / $perPart)
$made = @()
for ($i = 0; ($i -lt $total) -and (-not $UnpackerOnly); $i++) {
    $take = [math]::Min($perPart, $payload.Count - $i * $perPart)
    $lines = New-Header -No ($i + 1) -Of $total
    $lines += $payload.GetRange($i * $perPart, $take)
    $lines += ('<<< КОНЕЦ ЧАСТИ ' + ($i + 1) + ' ИЗ ' + $total + ' >>>')
    $f = Join-Path $OutDir ('LineProfiler-mail-{0:d2}.txt' -f ($i + 1))
    [IO.File]::WriteAllLines($f, $lines, [Text.UTF8Encoding]::new($true))
    $made += $f
}
Remove-Item $zip -Force


# --- Распаковщик на C/AL -------------------------------------------------------
# Второй путь на случай, когда PowerShell на сервере запускать нельзя, а C/SIDE есть.
# Объект СПЛОШЬ ЛАТИНСКИЙ - и код, и подписи: он едет телом письма и сохраняется кем
# угодно в какой угодно кодировке, а на латинице ни одна кодировка не врёт.
# Приём WriteAllBytes(Convert.FromBase64String(...)) взят из штатных объектов NAV
# (CU 1287, CU 1410) - значит платформа его точно принимает.
# Распаковщик на C/AL читает части С ДИСКА СЕРВЕРА, а туда файл положить дают не
# везде - по умолчанию его не кладём, чтобы не предлагать заведомо неработающий путь.
$tpl = Join-Path $onsiteDir 'Unpacker.template.txt'
if (-not (Test-Path $tpl)) { throw "нет шаблона распаковщика: $tpl" }
$boot = [IO.File]::ReadAllText($tpl)
# В шаблоне стоят НАСТОЯЩИЕ даты, а не плейсхолдеры: так он сам остаётся валидным
# объектом C/SIDE и импортируется наравне с остальными объектами профайлера.
# Размеры в объект больше не вшиваются: он обязан подходить к любому набору частей.
$boot = $boot -replace 'Date=\d\d\.\d\d\.\d\d;', ('Date=' + (Get-Date -Format 'dd.MM.yy') + ';')
$boot = $boot -replace 'DD \d\d\.\d\d\.\d{4} LineProfiler', ('DD ' + (Get-Date -Format 'dd.MM.yyyy') + ' LineProfiler')

# ДВА письма, а не одно: C/SIDE не примет файл, у которого перед OBJECT стоит хоть
# строка текста, а инструкцию по-русски в объект не вложишь. Объект уходит отдельным
# письмом и сохраняется КАК ЕСТЬ - латиница, без BOM, ничего лишнего сверху.
$bootLines = @(
    '=== ПОСТРОЧНЫЙ ПРОФАЙЛЕР C/AL - РАСПАКОВЩИК, ИНСТРУКЦИЯ (ЧАСТЬ 0A) ===',
    '',
    'Нужен, только если PowerShell на сервере запускать нельзя. Распаковщик - объект',
    'NAV: он сам соберёт архив из частей 1..N. Сам объект пришёл ОТДЕЛЬНЫМ письмом',
    '(ЧАСТЬ 0B) - его тело нельзя ничем дополнять, иначе C/SIDE не примет файл.',
    '',
    'ЧТО СДЕЛАТЬ',
    '  1. Тело письма ЧАСТЬ 0B сохранить как unpacker.txt - целиком и без правок.',
    '     Кодировка любая: там сплошная латиница, ломаться нечему.',
    '  2. C/SIDE -> File -> Import -> unpacker.txt, скомпилировать Codeunit 110207.',
    '  3. Части 1..N сохранить НА СЕРВЕРЕ NAV как C:\LineProfiler\part01.txt и далее.',
    '  4. Object Designer -> Codeunit 110207 -> Run. Он сверит размер и покажет итог.',
    '  5. Объект сохранит на ВАШУ машину один файл - LineProfilerPackage.zip.',
    '     Сборка приёмника лежит внутри него, в каталоге receiver.',
    '  6. Распаковать LineProfilerPackage.zip Проводником, дальше по README.md.',
    '  7. Codeunit 110207 после этого можно удалить.',
    '',
    ('Ожидаемый размер архива: ' + $bytes.Length + ' байт; свой объект покажет в конце.'),
    'Не сошлось - какая-то часть пришла обрезанной, и пакетом пользоваться нельзя.',
    'Размер объект намеренно не сверяет сам: это привязало бы его к одной рассылке.',
    '',
    'Файлы частей лежат на СЕРВЕРЕ: работа с файлами из C/AL идёт на стороне службы.',
    'Каталог ищется сам среди C:\LineProfiler, C:\Temp\LineProfiler,',
    'C:\ProgramData\LineProfiler, C:\Temp - в каком найдётся part01.txt, тот и берётся.'
)
$bootFile = Join-Path $OutDir 'LineProfiler-mail-00a-instructions.txt'
[IO.File]::WriteAllLines($bootFile, $bootLines, [Text.UTF8Encoding]::new($true))

# Объект - БЕЗ BOM: C/SIDE спотыкается о невидимые байты перед OBJECT.
$objFile = Join-Path $OutDir 'LineProfiler-mail-00b-unpacker-object.txt'
[IO.File]::WriteAllText($objFile, $boot, [Text.UTF8Encoding]::new($false))
Step ('распаковщик C/AL: инструкция + объект (' + $boot.Length + ' знаков)')


Write-Host ''
Write-Host 'ИТОГ' -ForegroundColor Green
foreach ($f in $made) { Step ("{0}  {1} КБ" -f (Split-Path $f -Leaf), [math]::Round((Get-Item $f).Length / 1KB)) }
Step "сумма архива: $hash"
