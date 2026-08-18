#requires -Version 7
<#
.SYNOPSIS
    Валидатор C/AL task-файла: кодировка, переводы строк, порядок объектов,
    обязательные свойства OBJECT-PROPERTIES, парность тегов изменений и
    запрещённая типографика.

.DESCRIPTION
    Проверяет один .txt с объектами задачи (классический текстовый экспорт C/SIDE):
      (a) валидный UTF-8 без BOM; отсутствие одиночных LF (все переводы строк CRLF);
      (b) порядок OBJECT-заголовков по типам:
          Table -> Page -> Report -> Codeunit -> XMLport -> Query -> MenuSuite
          (внутри одного типа порядок ID не важен);
      (c) у каждого OBJECT в блоке OBJECT-PROPERTIES:
          Date= в формате ДД.ММ.ГГ (двузначный год); Modified=Yes.
          Version List НЕ проверяется на код задачи: в этом проекте код задачи
          в список версий не дописывается (там только метки клиента) — проверяется
          лишь наличие строки, и то предупреждением;
      (d) парность тегов изменений: число строк «// <код> ... >>» равно числу
          «// <код> ... <<»; при последовательном проходе вложенность не уходит в минус;
      (e) запрещённая типографика: — – … « » „ “ ” и NBSP (печатается строка и символ);
      (f) TextConst-эвристика (warning): в TextConst-строке внутри одинарных кавычек
          после RUS=/ENU= встречается ещё один '=' вне удвоенных кавычек "";
      (g) чужие тег-комментарии (других задач/разработчиков) ВНУТРИ наших блоков
          >>..<< — при копировании чужого кода как шаблона их комментарии убираются.
          Чужие теги ВНЕ наших блоков не проверяются: в унаследованном коде объекта
          они легитимны и должны сохраняться.
          ВНИМАНИЕ: проверка (g) ОТКЛЮЧЕНА - список префиксов чужих тегов
          зависит от базы и не заполнен (см. $foreignTagPrefixes ниже).

    Вывод: список проблем с номерами строк и итоговая сводка.
    Код возврата: 1 если есть ошибки (errors); warnings не валят (код 0).

    ЭТОТ скрипт проверяет только МЕТА-конвенции: оформление правок, кодировку,
    порядок объектов. Качество самого кода C/AL (отступы, неиспользуемые
    переменные, TextConstant/CaptionML и прочее) он не смотрит - для этого нужен
    анализатор C/AL, если он есть в вашем контуре.


.PARAMETER TaskFile
    Путь к C/AL task-файлу.

.PARAMETER TaskCode
    Код проекта, напр. LineProfiler — используется для проверки парности тегов
    изменений и чужих тегов (в Version List код задачи НЕ пишется).
    По умолчанию — имя файла без расширения.

.EXAMPLE
    pwsh scripts/Test-ObjectFile.ps1 LineProfiler.txt
    Валидирует файл, код берётся из имени файла (LineProfiler).

.EXAMPLE
    pwsh scripts/Test-ObjectFile.ps1 LineProfiler.txt -TaskCode LineProfiler
    Явно заданный код задачи.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TaskFile,
    [string]$TaskCode
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $TaskFile -PathType Leaf)) {
    Write-Host "ОШИБКА: файл не найден: $TaskFile" -ForegroundColor Red
    exit 1
}
$src = (Resolve-Path -LiteralPath $TaskFile).Path
if ([string]::IsNullOrWhiteSpace($TaskCode)) {
    $TaskCode = [System.IO.Path]::GetFileNameWithoutExtension($src)
}

$errors   = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
function AddError([int]$lineNo, [string]$msg) {
    if ($lineNo -gt 0) { $script:errors.Add(("  [строка {0}] {1}" -f $lineNo, $msg)) }
    else               { $script:errors.Add(("  {0}" -f $msg)) }
}
function AddWarn([int]$lineNo, [string]$msg) {
    if ($lineNo -gt 0) { $script:warnings.Add(("  [строка {0}] {1}" -f $lineNo, $msg)) }
    else               { $script:warnings.Add(("  {0}" -f $msg)) }
}

# --- (a) Кодировка/BOM/LF ---
$bytes = [System.IO.File]::ReadAllBytes($src)

# BOM (UTF-8) проверка
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    AddError 0 "Файл начинается с UTF-8 BOM (EF BB BF) — требуется UTF-8 без BOM."
}

# Валидность UTF-8 (строгое декодирование)
$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)  # без BOM, throwOnInvalidBytes=true
$text = $null
try {
    $text = $strictUtf8.GetString($bytes)
    # GetString с throwOnInvalid может не бросить на некоторых последовательностях; декодер надёжнее:
    $dec = [System.Text.Encoding]::GetEncoding('utf-8',
            [System.Text.EncoderExceptionFallback]::new(),
            [System.Text.DecoderExceptionFallback]::new())
    [void]$dec.GetString($bytes)
} catch {
    AddError 0 ("Файл не является валидным UTF-8: {0}" -f $_.Exception.Message)
    $text = $null
}

if ($null -ne $text) {
    # Одиночные LF: позиция \n, не предварённая \r
    for ($i = 0; $i -lt $text.Length; $i++) {
        if ($text[$i] -eq "`n") {
            if ($i -eq 0 -or $text[$i - 1] -ne "`r") {
                # номер строки = число CRLF/LF до текущей позиции + 1
                $pre = $text.Substring(0, $i)
                $ln  = ([regex]::Matches($pre, "`n")).Count + 1
                AddError $ln "Одиночный LF (перевод строки не CRLF)."
            }
        }
    }
}

# Если файл не читается — дальше анализировать нечего
if ($null -eq $text) {
    Write-Host "=== Проблемы ===" -ForegroundColor Red
    foreach ($e in $errors) { Write-Host $e -ForegroundColor Red }
    Write-Host ""
    Write-Host ("ИТОГ: ошибок {0}, предупреждений {1}" -f $errors.Count, $warnings.Count) -ForegroundColor Red
    exit 1
}

# Разбиваем на строки (по CRLF и одиночному LF — чтобы нумерация совпадала с редактором)
$lines = [regex]::Split($text, "`r`n|`n")

# --- Типографские символы (для проверки e) ---
$forbidden = @{
    [char]0x2014 = 'EM DASH (—)'
    [char]0x2013 = 'EN DASH (–)'
    [char]0x2026 = 'ELLIPSIS (…)'
    [char]0x00AB = 'LEFT GUILLEMET («)'
    [char]0x00BB = 'RIGHT GUILLEMET (»)'
    [char]0x201E = 'DOUBLE LOW-9 („)'
    [char]0x201C = 'LEFT DQUOTE (“)'
    [char]0x201D = 'RIGHT DQUOTE (”)'
    [char]0x00A0 = 'NBSP'
}

# Регэкспы
$reHeader  = '^OBJECT\s+(Table|Page|Report|Codeunit|XMLport|Query|MenuSuite)\s+(\d+)\s+(.*?)\s*$'
$typeOrder = @{ 'Table'=1; 'Page'=2; 'Report'=3; 'Codeunit'=4; 'XMLport'=5; 'Query'=6; 'MenuSuite'=7 }

# Экранированный код задачи для построения шаблонов тегов
$tcEsc = [regex]::Escape($TaskCode)

# (g) Шаблон чужих тег-комментариев (тегов других задач/разработчиков в baseline).
#     Список префиксов ПУСТ - проверка отключена: семейства чужих тегов зависят
#     от конкретной базы и заранее неизвестны. Чтобы включить - собрать grep'ом
#     по выгрузке объектов фактически встречающиеся префиксы (напр. 'ABC','XYZ')
#     и перечислить их здесь — проверка включится сама.
#     Пустой список оставлен намеренно: чужой (непроверенный) шаблон дал бы
#     ложно-зелёную проверку.
$foreignTagPrefixes = @()
$reForeignTag = if ($foreignTagPrefixes.Count -gt 0) {
    '//.*?\b(?:[A-Z]{2,4}\.)?(?:' + (($foreignTagPrefixes | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')-?\d{3,6}\b'
} else {
    $null
}
if ($null -eq $reForeignTag) {
    AddWarn 0 'Проверка чужих тегов отключена: префиксы не перечислены (заполнить $foreignTagPrefixes в scripts/Test-ObjectFile.ps1).'
}

# Подсчёт по объектам + порядок типов
$objects = [System.Collections.Generic.List[object]]::new()
$prevRank = 0
$objCount = 0

# Состояние блока OBJECT-PROPERTIES для текущего объекта
$inObjProps = $false
$curObj = $null

$nestDepth   = 0
$openCount   = 0
$closeCount  = 0
$minNest     = 0

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line   = $lines[$i]
    $lineNo = $i + 1

    # (e) Запрещённая типографика
    foreach ($ch in $forbidden.Keys) {
        if ($line.IndexOf([char]$ch) -ge 0) {
            $code = [int][char]$ch
            AddError $lineNo ("Запрещённая типографика: {0} U+{1:X4}. Строка: {2}" -f $forbidden[$ch], $code, $line.Trim())
        }
    }

    # OBJECT-заголовок
    if ($line -match $reHeader) {
        $objCount++
        $type = $Matches[1]
        $id   = $Matches[2]
        $rank = $typeOrder[$type]
        if ($rank -lt $prevRank) {
            AddError $lineNo ("Нарушен порядок типов: {0} {1} идёт после типа с большим приоритетом (требуется Table->Page->Report->Codeunit->XMLport->Query->MenuSuite)." -f $type, $id)
        }
        $prevRank = $rank
        $curObj = [PSCustomObject]@{
            Type='OK'; HeaderLine=$lineNo; ObjType=$type; ObjId=$id
            HasVersionList=$false; VLLine=0
            HasDate=$false; DateOk=$false; DateLine=0
            HasModified=$false; ModifiedYes=$false; ModLine=0
        }
        $objects.Add($curObj)
        $inObjProps = $false
        continue
    }

    # Вход/выход блока OBJECT-PROPERTIES
    if ($line.Trim() -eq 'OBJECT-PROPERTIES') { $inObjProps = $true; continue }
    if ($inObjProps -and $line.Trim() -eq 'PROPERTIES') { $inObjProps = $false }

    if ($inObjProps -and $null -ne $curObj) {
        $t = $line.Trim()
        if ($t -match '^Version List=') {
            $curObj.HasVersionList = $true
            $curObj.VLLine = $lineNo
        }
        elseif ($t -match '^Date=(.*?);?\s*$') {
            $curObj.HasDate = $true
            $curObj.DateLine = $lineNo
            $dval = $Matches[1].TrimEnd(';').Trim()
            if ($dval -match '^\d{2}\.\d{2}\.\d{2}$') { $curObj.DateOk = $true }
        }
        elseif ($t -match '^Modified=(.*?);?\s*$') {
            $curObj.HasModified = $true
            $curObj.ModLine = $lineNo
            if ($Matches[1].TrimEnd(';').Trim() -eq 'Yes') { $curObj.ModifiedYes = $true }
        }
    }

    # (d) Парность тегов изменений «// <Инициалы> <Дата> <код> >>» / «... <<»
    #     Порядок тега — стиль baseline (код в конце), поэтому код ищется В ЛЮБОМ месте
    #     комментария, а не сразу после //. Тег должен содержать код задачи и заканчиваться маркером.
    if ($line -match ("//.*\b{0}\b.*>>\s*$" -f $tcEsc)) {
        $openCount++; $nestDepth++
    }
    elseif ($line -match ("//.*\b{0}\b.*<<\s*$" -f $tcEsc)) {
        $closeCount++; $nestDepth--
        if ($nestDepth -lt $minNest) { $minNest = $nestDepth }
        if ($nestDepth -lt 0) {
            AddError $lineNo "Закрывающий тег << без соответствующего открывающего >> (вложенность ушла в минус)."
        }
    }
    # (g) Чужой тег-комментарий внутри нашего блока >>..<< — скопированный шаблон с чужим комментарием
    #     ($reForeignTag = $null -> проверка отключена; пустой шаблон совпал бы с любой строкой)
    elseif ($nestDepth -gt 0 -and $null -ne $reForeignTag -and $line -match $reForeignTag -and $line -notmatch $tcEsc) {
        AddError $lineNo ("Чужой тег-комментарий внутри блока {0} (при копировании чужого кода комментарии убирать): {1}" -f $TaskCode, $line.Trim())
    }

    # (f) TextConst-эвристика (warning)
    if ($line -match 'TextConst') {
        # Берём содержимое первой одинарной строки '...'
        $m = [regex]::Match($line, "TextConst\s+'(.*)'")
        if ($m.Success) {
            $content = $m.Groups[1].Value
            # Убираем удвоенные кавычки "" (экранированные внутри C/AL TextConst)
            $stripped = $content -replace '""', ''
            # Ищем сегменты после RUS=/ENU= и проверяем, есть ли ещё '=' до ';' или конца
            $segs = [regex]::Matches($stripped, '(?:RUS=|ENU=)([^;]*)')
            foreach ($seg in $segs) {
                $segVal = $seg.Groups[1].Value
                if ($segVal.Contains('=')) {
                    AddWarn $lineNo ("TextConst: после RUS=/ENU= встречается ещё один '=' вне удвоенных кавычек. Сегмент: {0}" -f $segVal.Trim())
                    break
                }
            }
        }
    }
}

# (c) Проверка свойств по каждому объекту
foreach ($o in $objects) {
    $hdr = ("OBJECT {0} {1}" -f $o.ObjType, $o.ObjId)
    # Version List в этом проекте НЕ трогаем: код задачи туда НЕ дописывается.
    # В списке версий живут только метки клиента (Marking/EG/HS/Mercury и пр.) —
    # их ставит релиз-процесс клиента, а не мы. Поэтому проверка только на наличие
    # строки, и то предупреждением: у части объектов baseline она пуста/отсутствует.
    if (-not $o.HasVersionList) {
        AddWarn $o.HeaderLine ("{0}: нет строки Version List= в OBJECT-PROPERTIES — сверить с baseline." -f $hdr)
    }
    if (-not $o.HasDate) {
        AddError $o.HeaderLine ("{0}: отсутствует строка Date= в OBJECT-PROPERTIES." -f $hdr)
    } elseif (-not $o.DateOk) {
        AddError $o.DateLine ("{0}: Date= не в формате ДД.ММ.ГГ (двузначный год)." -f $hdr)
    }
    if (-not $o.HasModified) {
        AddError $o.HeaderLine ("{0}: отсутствует строка Modified= в OBJECT-PROPERTIES." -f $hdr)
    } elseif (-not $o.ModifiedYes) {
        AddError $o.ModLine ("{0}: Modified не равно Yes." -f $hdr)
    }
}

# (d) Итоговая парность
if ($openCount -ne $closeCount) {
    AddError 0 ("Непарные теги изменений: открывающих >> = {0}, закрывающих << = {1} (для кода '{2}')." -f $openCount, $closeCount, $TaskCode)
}
if ($nestDepth -gt 0) {
    AddError 0 ("Незакрытые теги изменений: на конце файла вложенность = {0}." -f $nestDepth)
}

# --- Вывод ---
Write-Host ("Файл:        {0}" -f $src)
Write-Host ("Код задачи:  {0}" -f $TaskCode)
Write-Host ("Объектов:    {0}" -f $objCount)
Write-Host ("Теги >>/<<:  {0} / {1}" -f $openCount, $closeCount)
Write-Host ""

if ($errors.Count -gt 0) {
    Write-Host "=== ОШИБКИ (errors) ===" -ForegroundColor Red
    foreach ($e in $errors) { Write-Host $e -ForegroundColor Red }
    Write-Host ""
}
if ($warnings.Count -gt 0) {
    Write-Host "=== ПРЕДУПРЕЖДЕНИЯ (warnings) ===" -ForegroundColor Yellow
    foreach ($w in $warnings) { Write-Host $w -ForegroundColor Yellow }
    Write-Host ""
}
if ($errors.Count -eq 0 -and $warnings.Count -eq 0) {
    Write-Host "Проблем не найдено." -ForegroundColor Green
}

Write-Host ("ИТОГ: ошибок {0}, предупреждений {1}" -f $errors.Count, $warnings.Count) -ForegroundColor $(if ($errors.Count -gt 0) {'Red'} else {'Green'})

if ($errors.Count -gt 0) { exit 1 } else { exit 0 }
