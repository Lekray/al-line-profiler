#requires -Version 7
<#
.SYNOPSIS
    Готовит C/AL task-файл к импорту в C/SIDE: заменяет типографику на ASCII,
    нормализует переводы строк в CRLF и пишет копию в кодировке cp866.

.DESCRIPTION
    Читает указанный .txt (UTF-8 без BOM), НЕ изменяя исходный файл. Выполняет:
      * замену типографских символов на ASCII-аналоги
        (длинное/короткое тире -> '-', многоточие -> '...',
         кавычки-ёлочки и «лапки» -> '"', одинарные типографские -> "'",
         неразрывный пробел U+00A0 -> обычный пробел);
      * нормализацию переводов строк в CRLF;
      * проверку, что ВСЕ символы кодируются в cp866 (DOS Cyrillic, CP 866).
    Если после замен остаются символы, непреобразуемые в cp866, скрипт печатает
    их с номерами строк, НЕ пишет выходной файл и завершается с кодом 1.
    Иначе пишет «<папка>\<имя-без-расширения>.cp866.txt» в кодировке cp866 (CRLF).

    Исходный файл НИКОГДА не модифицируется.

    Кодировка cp866 принята по опыту C/SIDE: cp1251 и UTF-8 ломают кириллицу при
    импорте. На незнакомой среде её стоит подтвердить пробным импортом.

.PARAMETER TaskFile
    Путь к исходному C/AL task-файлу (UTF-8 без BOM).

.EXAMPLE
    pwsh scripts/convert-for-cside.ps1 LineProfiler.txt
    Создаёт LineProfiler.cp866.txt в кодировке cp866.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TaskFile
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $TaskFile -PathType Leaf)) {
    Write-Host "ОШИБКА: файл не найден: $TaskFile" -ForegroundColor Red
    exit 1
}

$src = (Resolve-Path -LiteralPath $TaskFile).Path

# Читаем как UTF-8 без BOM (raw, без преобразования переводов строк)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$text = [System.IO.File]::ReadAllText($src, $utf8NoBom)

# --- Замена типографики на ASCII ---
# Используем код-поинты, чтобы не зависеть от кодировки самого скрипта.
$replacements = [ordered]@{
    [char]0x2014 = '-'    # — em dash
    [char]0x2013 = '-'    # – en dash
    [char]0x2026 = '...'  # … horizontal ellipsis
    [char]0x00AB = '"'    # « left guillemet
    [char]0x00BB = '"'    # » right guillemet
    [char]0x201E = '"'    # „ double low-9
    [char]0x201C = '"'    # “ left double quote
    [char]0x201D = '"'    # ” right double quote
    [char]0x2018 = "'"    # ‘ left single quote
    [char]0x2019 = "'"    # ’ right single quote
    [char]0x00A0 = ' '    # NBSP
}
foreach ($from in $replacements.Keys) {
    $text = $text.Replace([string]$from, $replacements[$from])
}

# --- Нормализация переводов строк в CRLF ---
# Сначала всё к LF, затем LF -> CRLF (чтобы не сдвоить \r).
$text = $text -replace "`r`n", "`n"
$text = $text -replace "`r", "`n"
$text = $text -replace "`n", "`r`n"

# --- Проверка кодируемости в cp866 посимвольно ---
$cp866 = [System.Text.Encoding]::GetEncoding(866)

$lines = [regex]::Split($text, "`r`n")
$problems = [System.Collections.Generic.List[string]]::new()
for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    for ($j = 0; $j -lt $line.Length; $j++) {
        $ch = $line[$j]
        # Кодируем символ и декодируем обратно: при потере (замене на '?') символ непреобразуем.
        $bytes = $cp866.GetBytes([string]$ch)
        $round = $cp866.GetString($bytes)
        if ($round -ne [string]$ch) {
            $code = [int][char]$ch
            $problems.Add(("  строка {0}, позиция {1}: U+{2:X4} '{3}'" -f ($i + 1), ($j + 1), $code, $ch))
        }
    }
}

if ($problems.Count -gt 0) {
    Write-Host "ОШИБКА: найдены символы, непреобразуемые в cp866 (после замены типографики):" -ForegroundColor Red
    foreach ($p in $problems) { Write-Host $p -ForegroundColor Yellow }
    Write-Host ""
    Write-Host ("Всего проблемных символов: {0}. Выходной файл НЕ записан." -f $problems.Count) -ForegroundColor Red
    exit 1
}

# --- Запись выходного файла в cp866 ---
$dir  = Split-Path -Parent $src
$name = [System.IO.Path]::GetFileNameWithoutExtension($src)
$outPath = Join-Path $dir ($name + '.cp866.txt')

$bytesOut = $cp866.GetBytes($text)
[System.IO.File]::WriteAllBytes($outPath, $bytesOut)

Write-Host "OK: типографика заменена, переводы строк CRLF, все символы кодируются в cp866." -ForegroundColor Green
Write-Host ("Записан файл (cp866): {0}" -f $outPath)
Write-Host ("Исходный файл не изменён: {0}" -f $src)
exit 0
