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
          Работает в двух режимах, и первый включён всегда:
            * по ФОРМЕ (предупреждение): комментарий внутри нашего блока начинается
              с тега вида ABC-12345 — две-шесть заглавных латинских букв и четыре-шесть
              цифр подряд. Знать конкретную базу для этого не нужно, поэтому режим
              работает на чистой копии и в чужом контуре;
            * по СПИСКУ (ошибка): семейства префиксов заданы явно — ключом
              -ForeignTagPrefix или переменной окружения LP_FOREIGN_TAGS (через запятую,
              точку с запятой или пробел). Захардкоженного списка здесь нет намеренно:
              репозиторий публичный, а префиксы задач принадлежат установке.

      (h) согласованность с двоичным контейнером: если рядом лежит <имя>.fob,
          сверяются его оглавление (состав объектов, имена, даты, списки версий) и
          происхождение — хэш текста, с которого .fob снят, из файла fob-origin.txt
          рядом. Оглавление ловит подмену состава, происхождение — отставание: дата
          в OBJECT-PROPERTIES ставится руками, и правка кода в тот же день оглавление
          не меняет вовсе. Проверено на живом расхождении: .fob отставал от текста на
          восемь коммитов, а состав совпадал полностью. Отключается ключом -NoFob.

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

.PARAMETER ForeignTagPrefix
    Префиксы тегов чужих задач, напр. ABC, XYZ. Найденные внутри наших блоков
    >>..<< — ошибка, а не предупреждение. Если не задан, берётся из переменной
    окружения LP_FOREIGN_TAGS (разделители: запятая, точка с запятой, пробел).
    В коде списка нет и не будет: репозиторий публичный.

.PARAMETER NoFob
    Не сверять с двоичным контейнером <имя>.fob, даже если он лежит рядом.

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
    [string]$TaskCode,
    [string[]]$ForeignTagPrefix,
    [switch]$NoFob
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

# (g) Чужие тег-комментарии внутри наших блоков >>..<<.
#
#     Проверка по ФОРМЕ работает всегда и знания конкретной базы не требует: тег
#     задачи выглядит как ABC-12345 - две-шесть заглавных латинских букв и следом
#     четыре-шесть цифр. Якорь на НАЧАЛО комментария не для строгости, а против
#     ложных срабатываний: 'CP1251', 'ISO8601' и 'RFC3339' в середине фразы имеют
#     ту же форму, и без якоря каждое второе пояснение про кодировки становилось бы
#     находкой. Тег же по соглашению стоит первым словом комментария.
#     Три цифры не берутся намеренно: 'SHA256', 'CP866' и 'NAVW111' - не теги.
$reTagShape = '//\s*(?:[A-Z]{2,4}\.)?[A-Z]{2,6}-?\d{4,6}(?![A-Za-z0-9])'

#     Проверка по СПИСКУ поднимает находку до ошибки, но список приходит СНАРУЖИ -
#     ключом или переменной окружения, как LP_DATABASE и LP_BASELINE_DIR. В коде его
#     нет по двум причинам: репозиторий публичный (префиксы задач принадлежат
#     установке), и захардкоженный чужой список дал бы ложно-зелёную проверку.
$foreignTagPrefixes = @()
if ($ForeignTagPrefix) { $foreignTagPrefixes = @($ForeignTagPrefix) }
elseif ($env:LP_FOREIGN_TAGS) {
    $foreignTagPrefixes = @($env:LP_FOREIGN_TAGS -split '[,;\s]+' | Where-Object { $_ })
}
$reForeignTag = $null
if (@($foreignTagPrefixes).Count -gt 0) {
    $reForeignTag = '//\s*(?:[A-Z]{2,4}\.)?(?:' +
        ((@($foreignTagPrefixes) | ForEach-Object { [regex]::Escape($_) }) -join '|') +
        ')-?\d{3,6}(?![A-Za-z0-9])'
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
    # (g) Чужой тег-комментарий внутри нашего блока >>..<< — скопированный шаблон
    #     с чужим комментарием. Известный префикс — ошибка, просто тег-по-форме —
    #     предупреждение: форма угадывает, список знает.
    elseif ($nestDepth -gt 0 -and $line -notmatch $tcEsc -and
            $null -ne $reForeignTag -and $line -match $reForeignTag) {
        AddError $lineNo ("Чужой тег-комментарий внутри блока {0} (при копировании чужого кода комментарии убирать): {1}" -f $TaskCode, $line.Trim())
    }
    elseif ($nestDepth -gt 0 -and $line -notmatch $tcEsc -and $line -match $reTagShape) {
        AddWarn $lineNo ("Комментарий внутри блока {0} начинается с тега чужой задачи (при копировании чужого кода комментарии убирать): {1}" -f $TaskCode, $line.Trim())
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

# --- (h) Согласованность с двоичным контейнером .fob -------------------------
# Проверка включается сама, если .fob лежит рядом. Две независимые стороны:
# оглавление отвечает «те ли это объекты», fob-origin.txt — «из этого ли текста
# он собран». Второе без первого бесполезно, первое без второго — ложно-зелено.
$fobPath = [System.IO.Path]::ChangeExtension($src, '.fob')
if (-not $NoFob -and (Test-Path -LiteralPath $fobPath -PathType Leaf)) {
    . (Join-Path $PSScriptRoot 'Lib-CSideFob.ps1')
    $fobName = [System.IO.Path]::GetFileName($fobPath)
    $txtName = [System.IO.Path]::GetFileName($src)

    $fobDir = @()
    try { $fobDir = @(Get-CSideFobDirectory -Path $fobPath) }
    catch { AddWarn 0 ("Оглавление {0} не прочиталось ({1}) — сверка состава пропущена." -f $fobName, $_.Exception.Message) }

    if ($fobDir.Count -eq 0) {
        AddWarn 0 ("В {0} не нашлось оглавления объектов — сверка состава пропущена." -f $fobName)
    }
    else {
        $txtDir = @(Get-CSideTextDirectory -Path $src)
        foreach ($d in @(Compare-CSideDirectory -Fob $fobDir -Text $txtDir)) {
            AddError 0 ("{0} расходится с текстом: {1}" -f $fobName, $d)
        }
    }

    # Происхождение. Хэш текста в манифесте — это декларация «фоб снят с ЭТОГО
    # текста», и она проверяемая: любая последующая правка её ломает.
    $originPath = Join-Path (Split-Path -Parent $src) 'fob-origin.txt'
    if (-not (Test-Path -LiteralPath $originPath -PathType Leaf)) {
        AddWarn 0 ("Рядом с {0} нет fob-origin.txt: из какого текста снят контейнер — неизвестно." -f $fobName)
    }
    else {
        $origin = @{}
        foreach ($line in [System.IO.File]::ReadAllLines($originPath, [System.Text.Encoding]::UTF8)) {
            $m = [regex]::Match($line, '^\s*([a-z]+)\s+(\S+)\s*$')
            if ($m.Success) { $origin[$m.Groups[1].Value] = $m.Groups[2].Value.ToLowerInvariant() }
        }
        $shaTxt = Get-CSideSha256 $src
        $shaFob = Get-CSideSha256 $fobPath

        if (-not $origin.ContainsKey('txt') -or -not $origin.ContainsKey('fob')) {
            AddWarn 0 'В fob-origin.txt нет строк txt и fob — происхождение контейнера не объявлено.'
        }
        else {
            if ($origin['fob'] -ne $shaFob) {
                AddError 0 ("{0} подменён после объявления: в fob-origin.txt {1}, на диске {2}." -f
                            $fobName, $origin['fob'].Substring(0, 16), $shaFob.Substring(0, 16))
            }
            if ($origin['txt'] -ne $shaTxt) {
                AddError 0 ("{0} снят с другой версии текста: в fob-origin.txt хэш {1}, у {2} сейчас {3}." -f
                            $fobName, $origin['txt'].Substring(0, 16), $txtName, $shaTxt.Substring(0, 16))
                AddError 0 ("Импорт {0} поставит код без последних правок текста. Перевыгрузить контейнер с базы, в которую импортирован текущий {1}, и обновить манифест: pwsh scripts/Update-FobOrigin.ps1" -f
                            $fobName, $txtName)
            }
        }
    }
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
