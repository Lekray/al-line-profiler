#requires -Version 7
<#
.SYNOPSIS
    Превращает любой файл в текст base64, пригодный для отправки телом письма.

.DESCRIPTION
    Нужно там, где почта контура режет вложения, а положить файл на сервер нельзя.

    Строки полезной нагрузки помечены в НАЧАЛЕ знаком '|'. Отбор идёт по метке, а не
    по длине: у последней строки каждой части длина своя, и отбор по длине её терял -
    архив собирался битым, а выглядело это как удачная отправка.

    Заголовок написан по-русски, полезная нагрузка - чистая латиница. Поэтому что бы
    почта ни сделала с кодировкой заголовка, сам файл восстановится побайтно.

    Если нагрузка не влезает в одно письмо, она режется на части. Части собираются в
    ЛЮБОМ порядке - каждая несёт свой номер, а строки внутри не разрываются.

.PARAMETER Path
    Файл, который надо упаковать.

.PARAMETER PartChars
    Знаков полезной нагрузки в одной части. По умолчанию 44800 (кратно 64) - около 46 КБ
    тела письма. Если всё влезает в одну часть, деления не будет.

.EXAMPLE
    pwsh scripts/ConvertTo-MailBase64.ps1 dist/AlLineProfiler.dll
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $Path,
    [string] $OutDir,
    [int]    $PartChars = 44800
)

$ErrorActionPreference = 'Stop'

$file = Get-Item $Path
$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$bytes = [IO.File]::ReadAllBytes($file.FullName)
$b64   = [Convert]::ToBase64String($bytes)
$hash  = (Get-FileHash $file.FullName -Algorithm SHA256).Hash.ToLower()

Write-Host 'Упаковка файла в текст письма' -ForegroundColor Cyan
Write-Host ("  {0}: {1:N0} байт -> {2:N0} знаков base64" -f $file.Name, $bytes.Length, $b64.Length)

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
        'Размер  : {0:N0} байт' -f $bytes.Length
        'SHA-256 : ' + $hash
        ''
        'Строки нагрузки начинаются со знака | в первой позиции. Этот заголовок вырезать'
        'не нужно - сборщик берёт только помеченные строки, остальное пропускает.'
        ''
        'КАК ВЕРНУТЬ ФАЙЛ (PowerShell, все части в одном каталоге):'
        ''
        '  $b64 = (Get-ChildItem *.txt | Sort-Object Name | Get-Content |'
        '          Where-Object { $_.StartsWith(''|'') } | ForEach-Object { $_.Substring(1) }) -join '''''
        ('  [IO.File]::WriteAllBytes(''{0}'', [Convert]::FromBase64String($b64))' -f $file.Name)
        ('  (Get-FileHash ''{0}'' -Algorithm SHA256).Hash   # должно совпасть с SHA-256 выше' -f $file.Name)
        ''
        '--- ниже нагрузка, сохранять целиком ---'
        ''
    )
    $body = $rows.GetRange($p * $perPart, $take)
    $name = if ($total -eq 1) { '{0}.base64.txt' -f $file.BaseName }
            else { '{0}.base64.часть{1:d2}.txt' -f $file.BaseName, ($p + 1) }
    $out = Join-Path $OutDir $name
    [IO.File]::WriteAllLines($out, ($head + $body), [Text.UTF8Encoding]::new($true))
    $made += $out
    Write-Host ('  часть {0}/{1}: {2} ({3:N0} байт)' -f ($p + 1), $total, $name, (Get-Item $out).Length)
}

Write-Host ''
Write-Host ('ИТОГ: писем {0}, сумма файла {1}' -f $made.Count, $hash) -ForegroundColor Green
