#requires -Version 7
<#
.SYNOPSIS
    Превращает любой файл в текст base64, пригодный для отправки телом письма.

.DESCRIPTION
    Нужно там, где почта контура режет вложения, а положить файл на сервер нельзя.

    Строки полезной нагрузки помечены в НАЧАЛЕ знаком '|'. Отбор идёт по метке, а не
    по длине: у последней строки каждой части длина своя, и отбор по длине её терял -
    файл собирался битым, а выглядело это как удачная отправка.

    Заголовок написан по-русски, полезная нагрузка - чистая латиница. Поэтому что бы
    почта ни сделала с кодировкой заголовка, сам файл восстановится побайтно.

    Если нагрузка не влезает в одно письмо, она режется на части. Части собираются в
    ЛЮБОМ порядке - каждая несёт свой номер, а строки внутри не разрываются.

.PARAMETER Path
    Файл, который надо упаковать.

.PARAMETER Zip
    Завернуть файл в архив ПЕРЕД кодированием. Нужно, если собирать письмо будет
    распаковщик на C/AL (Codeunit 110207): он проверяет, что получился именно архив,
    и одиночный файл отвергает. Без этого ключа получается голый файл - его собирают
    командами из шапки письма.

.PARAMETER PartChars
    Знаков полезной нагрузки в одной части. По умолчанию 44800 (кратно 64) - около 46 КБ
    тела письма. Если всё влезает в одну часть, деления не будет.

.PARAMETER Plain
    Отдать ЧИСТЫЙ base64 одним файлом: ни меток, ни шапки, ни деления на части. Нужно,
    когда принимающая сторона разбирает файл СВОИМ разбирателем base64, а не нашим:
    чужой инструмент о метке '|' не знает и на ней спотыкается, а русская шапка для него
    просто мусор посреди документа. Пояснение и контрольная сумма уезжают отдельным
    файлом рядом, чтобы не попасть внутрь потока.

.PARAMETER LineWidth
    Длина строки в режиме -Plain. По умолчанию 76 - как в MIME, её принимают все
    разбиратели. 0 - выложить одной строкой без переносов вовсе.

.EXAMPLE
    pwsh scripts/ConvertTo-MailBase64.ps1 dist/AlLineProfiler.dll -Zip
    Письмо под распаковщик Codeunit 110207.

.EXAMPLE
    pwsh scripts/ConvertTo-MailBase64.ps1 dist/AlLineProfiler.dll
    Письмо, которое собирают вручную командами из его же шапки.

.EXAMPLE
    pwsh scripts/ConvertTo-MailBase64.ps1 out/LineProfiler-onsite-20260819.zip -Plain
    Чистый base64 под чужой разбиратель: один файл, ничего лишнего внутри.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $Path,
    [switch] $Zip,
    [switch] $Plain,
    [string] $OutDir,
    [int]    $PartChars = 44800,
    [int]    $LineWidth = 76
)

$ErrorActionPreference = 'Stop'

$file = Get-Item $Path
$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

Write-Host 'Упаковка файла в текст письма' -ForegroundColor Cyan

$srcHash = (Get-FileHash $file.FullName -Algorithm SHA256).Hash.ToLower()
$payloadFile = $file.FullName
$tmpZip = $null
if ($Zip) {
    # Распаковщик на C/AL узнаёт архив по сигнатуре и одиночный файл не принимает,
    # поэтому под него нагрузку заворачиваем.
    $tmpZip = Join-Path ([IO.Path]::GetTempPath()) ($file.BaseName + '-mail.zip')
    if (Test-Path $tmpZip) { [IO.File]::Delete($tmpZip) }
    Compress-Archive -Path $file.FullName -DestinationPath $tmpZip
    $payloadFile = $tmpZip
    Write-Host ("  завёрнут в архив: {0:N0} -> {1:N0} байт" -f $file.Length, (Get-Item $tmpZip).Length)
}

$bytes = [IO.File]::ReadAllBytes($payloadFile)
$b64   = [Convert]::ToBase64String($bytes)
$hash  = (Get-FileHash $payloadFile -Algorithm SHA256).Hash.ToLower()
Write-Host ("  {0}: {1:N0} байт -> {2:N0} знаков base64" -f $file.Name, $bytes.Length, $b64.Length)

# --- чистый base64: один файл, внутри ничего кроме нагрузки --------------------
if ($Plain) {
    # Ни меток, ни шапки, ни деления. Всё, что могло бы помочь человеку, лежит РЯДОМ,
    # а не внутри: чужой разбиратель base64 читает файл целиком, и любая наша строка
    # посреди потока для него - ошибка формата, а не подсказка.
    if ($LineWidth -lt 0) { throw 'LineWidth не может быть отрицательной' }
    if ($LineWidth -eq 0) {
        $plainText = $b64
    } else {
        $wrapped = [System.Collections.Generic.List[string]]::new()
        for ($i = 0; $i -lt $b64.Length; $i += $LineWidth) {
            $wrapped.Add($b64.Substring($i, [Math]::Min($LineWidth, $b64.Length - $i)))
        }
        $plainText = $wrapped -join "`r`n"
    }
    $suffix  = if ($Zip) { '.zip' } else { '' }
    $outB64  = Join-Path $OutDir ('{0}{1}.b64' -f $file.BaseName, $suffix)
    # Без BOM: три невидимых байта в начале - это три знака, которых base64 не знает,
    # и строгий разбиратель на них отказывает. Наш собственный тоже.
    [IO.File]::WriteAllText($outB64, $plainText, [Text.UTF8Encoding]::new($false))

    $note = @(
        'Что это'
        ''
        ('  Файл    : {0}' -f $file.Name)
        ('  Размер  : {0:N0} байт' -f $file.Length)
        ('  SHA-256 : {0}' -f $srcHash)
    )
    if ($Zip) {
        $note += @(
            ''
            ('  В base64 лежит АРХИВ с этим файлом внутри: {0:N0} байт, SHA-256 {1}.' -f $bytes.Length, $hash)
        )
    }
    $note += @(
        ''
        ('  Текст   : {0} ({1:N0} знаков base64)' -f (Split-Path $outB64 -Leaf), $b64.Length)
        ''
        'Внутри .b64 нет ничего, кроме самой нагрузки: ни заголовка, ни меток, ни BOM.'
        $(if ($LineWidth -eq 0) { 'Одна строка без переносов.' } else { ('Строки по {0} знаков, перенос CRLF - как в MIME.' -f $LineWidth) })
        ''
        'Проверить и разобрать без всяких инструментов:'
        ''
        ('  certutil -decode {0} out{1}' -f (Split-Path $outB64 -Leaf), $(if ($Zip) { '.zip' } else { [IO.Path]::GetExtension($file.Name) }))
        ('  certutil -hashfile out{0} SHA256' -f $(if ($Zip) { '.zip' } else { [IO.Path]::GetExtension($file.Name) }))
        ''
        ('Сумма должна совпасть с {0}.' -f $(if ($Zip) { $hash } else { $srcHash }))
        'Не совпала - файл дошёл не целиком, разбирать его нельзя.'
    )
    $outNote = $outB64 + '.txt'
    [IO.File]::WriteAllLines($outNote, $note, [Text.UTF8Encoding]::new($true))

    if ($tmpZip -and (Test-Path $tmpZip)) { [IO.File]::Delete($tmpZip) }
    Write-Host ('  {0}: {1:N0} байт' -f (Split-Path $outB64 -Leaf), (Get-Item $outB64).Length)
    Write-Host ('  {0}: пояснение и суммы' -f (Split-Path $outNote -Leaf))
    Write-Host ''
    Write-Host ('ИТОГ: чистый base64, сумма нагрузки {0}' -f $(if ($Zip) { $hash } else { $srcHash })) -ForegroundColor Green
    exit 0
}

# Ровно 64 знака в строке: так строка не разрывается на границе части.
$rows = [System.Collections.Generic.List[string]]::new()
for ($i = 0; $i -lt $b64.Length; $i += 64) {
    $rows.Add('|' + $b64.Substring($i, [Math]::Min(64, $b64.Length - $i)))
}
$perPart = [Math]::Floor($PartChars / 64)
$total   = [Math]::Max(1, [Math]::Ceiling($rows.Count / $perPart))

$made = @()
for ($p = 0; $p -lt $total; $p++) {
    $take = [Math]::Min($perPart, $rows.Count - $p * $perPart)
    $head = @(
        '=== ФАЙЛ В ВИДЕ ТЕКСТА - ЧАСТЬ {0} ИЗ {1} ===' -f ($p + 1), $total
        ''
        'Файл    : ' + $file.Name
        'Размер  : {0:N0} байт' -f $file.Length
        'SHA-256 : ' + $srcHash
    )
    if ($Zip) {
        $head += @(
            ''
            'Нагрузка - АРХИВ с этим файлом внутри ({0:N0} байт, SHA-256 {1}).' -f $bytes.Length, $hash
            'Так его принимает распаковщик на C/AL: одиночный файл он отвергает, потому что'
            'проверяет сигнатуру архива.'
            ''
            'КАК СОБРАТЬ В NAV: импортировать Codeunit 110207, запустить, указать части по'
            'очереди, на вопрос о следующей части нажать Отмену. Объект отдаст архив на вашу'
            'машину; распаковать его Проводником.'
        )
    }
    $head += @(
        ''
        'Строки нагрузки начинаются со знака | в первой позиции. Этот заголовок вырезать'
        'не нужно - сборщик берёт только помеченные строки, остальное пропускает.'
        ''
        'КАК СОБРАТЬ БЕЗ NAV (PowerShell, все части в одном каталоге):'
        ''
        '  $b64 = (Get-ChildItem *.txt | Sort-Object Name | Get-Content |'
        '          Where-Object { $_.StartsWith(''|'') } | ForEach-Object { $_.Substring(1) }) -join '''''
        ('  [IO.File]::WriteAllBytes(''{0}'', [Convert]::FromBase64String($b64))' -f (Split-Path $payloadFile -Leaf))
        ('  (Get-FileHash ''{0}'' -Algorithm SHA256).Hash   # должно совпасть с {1}' -f (Split-Path $payloadFile -Leaf), $(if ($Zip) { 'суммой архива выше' } else { 'SHA-256 выше' }))
        ''
        '--- ниже нагрузка, сохранять целиком ---'
        ''
        # Номер части ЛАТИНИЦЕЙ: шапка выше по-русски, а распаковщик на C/AL - сплошная
        # латиница, свой номер он оттуда прочитать не может. Без номера части склеиваются
        # в порядке выбора мышью, а base64, склеенный не в том порядке, остаётся ВАЛИДНЫМ
        # base64: раскодируется молча и даёт мусор.
        ('#PART {0:d2} OF {1:d2}' -f ($p + 1), $total)
        ''
    )
    $body = $rows.GetRange($p * $perPart, $take)
    $suffix = if ($Zip) { '.zip' } else { '' }
    $name = if ($total -eq 1) { '{0}{1}.base64.txt' -f $file.BaseName, $suffix }
            else { '{0}{1}.base64.часть{2:d2}.txt' -f $file.BaseName, $suffix, ($p + 1) }
    $out = Join-Path $OutDir $name
    [IO.File]::WriteAllLines($out, ($head + $body), [Text.UTF8Encoding]::new($true))
    $made += $out
    Write-Host ('  часть {0}/{1}: {2} ({3:N0} байт)' -f ($p + 1), $total, $name, (Get-Item $out).Length)
}

if ($tmpZip -and (Test-Path $tmpZip)) { [IO.File]::Delete($tmpZip) }

Write-Host ''
Write-Host ('ИТОГ: писем {0}, сумма исходного файла {1}' -f $made.Count, $srcHash) -ForegroundColor Green
