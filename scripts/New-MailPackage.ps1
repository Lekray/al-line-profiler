#requires -Version 7
<#
.SYNOPSIS
    Собирает пакет В ВИДЕ ТЕКСТА - для отправки телом письма.

.DESCRIPTION
    Почтовый фильтр контура удаляет вложения, поэтому пакет едет текстом: base64
    строками по 76 знаков, которые переживают любой перенос строк в почтовом клиенте.

    Двоичного в пакете нет вовсе: собранная сборка приёмника НЕ кладётся - её всё равно
    полагается пересобирать на месте (шаг 3), ссылкой на ТУ TraceEvent.dll, что лежит
    там. Исходник в UTF-8 тоже не кладётся: для импорта нужен cp866, а править объекты
    на месте удобнее уже в C/SIDE.

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
# Сборка вернулась в пакет. Она нужна на РАБОЧЕЙ СТАНЦИИ: положить её на сервер сможет
# либо установщик из C/AL (если у учётки службы есть права на Add-ins), либо человек с
# доступом к серверу - но в обоих случаях файл сперва должен доехать сюда.
$dll = Join-Path $taskDir 'bin\AlLineProfiler.dll'
if (-not (Test-Path $dll)) { throw "нет сборки приёмника: $dll (запустить Deploy-LineProfiler.ps1)" }

$zip = Join-Path $env:TEMP 'lineprofiler-mail.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Remove-Item $stage -Recurse -Force

$bytes = [IO.File]::ReadAllBytes($zip)
$hash  = (Get-FileHash $zip -Algorithm SHA256).Hash
# Сборка приёмника идёт ОТДЕЛЬНЫМ потоком под своей меткой, а не внутри архива:
# распаковщику на C/AL тогда не нужен разбор ZIP - одной зависимостью меньше, - и
# каждый файл сразу оказывается тем, чем должен быть.
$dllBytes = [IO.File]::ReadAllBytes($dll)
# Строки по 64 знака плюс метка - итого 65: заведомо короче даже 72, на которых
# переносит формат format=flowed. Перенос строки посреди нагрузки терял бы её молча:
# продолжение уехало бы на строку без метки, а такую строку отбор выбрасывает.
# Отбор ПО МЕТКЕ, а не по длине: у последней строки каждой части длина своя, и отбор
# по длине выбрасывал её - проверено обратной сборкой, сумма не сошлась.
$flat    = [Convert]::ToBase64String($bytes)
$flatDll = [Convert]::ToBase64String($dllBytes)
Step ("архив $([math]::Round($bytes.Length/1KB)) КБ + сборка $([math]::Round($dllBytes.Length/1KB)) КБ -> текст $([math]::Round(($flat.Length+$flatDll.Length)/1KB)) КБ")

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
        $h += ''
        $h += 'ЕСЛИ POWERSHELL НЕТ И НА РАБОЧЕЙ СТАНЦИИ'
        $h += '  Разбор base64 умеет штатный certutil, он есть в любой Windows:'
        $h += '     copy /b part01.txt+part02.txt+part03.txt+part04.txt+part05.txt all.txt'
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
    Get-ChildItem $OutDir -Filter 'LineProfiler-mail-[0-9][0-9].txt' -ErrorAction SilentlyContinue | Remove-Item -Force
}

# Сперва собираем ВСЕ строки нагрузки обоих потоков, потом режем по строкам: так ни
# одна строка не разрывается на границе части, а метка позволяет собирать в любом порядке.
$payload = New-Object System.Collections.Generic.List[string]
for ($j = 0; $j -lt $flat.Length; $j += 64) {
    $payload.Add('|' + $flat.Substring($j, [math]::Min(64, $flat.Length - $j)))
}
for ($j = 0; $j -lt $flatDll.Length; $j += 64) {
    $payload.Add('!' + $flatDll.Substring($j, [math]::Min(64, $flatDll.Length - $j)))
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
    '  5. Объект сохранит на ВАШУ машину два файла - имена ровно такие:',
    '        LineProfilerPackage.zip   (архив пакета)',
    '        AlLineProfiler.dll        (сборка приёмника)',
    '     Запомните каталог, который выберете в диалоге сохранения.',
    '  6. Распаковать LineProfilerPackage.zip Проводником, дальше по README.md.',
    '  7. Codeunit 110207 после этого можно удалить.',
    '',
    ('Ожидаемый размер архива: ' + $bytes.Length + ' байт. Не сойдётся - объект скажет'),
    'об этом ошибкой, и пакетом пользоваться нельзя: какая-то часть пришла обрезанной.',
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
