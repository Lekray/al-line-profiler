#requires -Version 7
<#
.SYNOPSIS
    Подменяет Documentation-блок во всех объектах текстового экспорта C/SIDE.

.DESCRIPTION
    В git у объектов один Documentation на всех - заголовок проекта: что за инструмент,
    автор, дата создания, ссылка на исходник. В базе, куда объекты уезжают, он ДРУГОЙ:
    там своя запись по номеру задачи трекера, по которой сопровождение потом и ищет,
    откуда объект взялся.

    Номер задачи принадлежит установке, а репозиторий публичный - в git его нет и быть
    не может. Поэтому текст приезжает СНАРУЖИ: ключом -Text или переменной окружения
    LP_DOC_TEXT, - и подставляется в пакет на выходе, не трогая исходник в репозитории.

    Блок узнаётся по форме: строка "BEGIN" с четырьмя пробелами отступа, следом строка
    с открывающей скобкой. У тела процедуры за BEGIN идёт код, а не скобка, поэтому
    спутать нельзя - но на всякий случай число заменённых блоков сверяется с числом
    объектов в файле, и расхождение роняет сборку.

.PARAMETER Path
    Файл текстового экспорта (UTF-8 без BOM, CRLF). Правится НА МЕСТЕ, поэтому
    подавать сюда рабочую копию из репозитория незачем - только копию в пакете.

.PARAMETER Text
    Что положить в блок. Несколько строк - через перевод строки. Если не задан ни ключ,
    ни LP_DOC_TEXT, файл не трогается вовсе и скрипт молча выходит: это законный случай
    сборки "как в git".

.EXAMPLE
    pwsh scripts/Set-Documentation.ps1 -Path stage/objects/LineProfiler.txt
    Берёт текст из LP_DOC_TEXT.

.EXAMPLE
    pwsh scripts/Set-Documentation.ps1 -Path o.txt -Text 'TASK-1 AB Object created'
    Тот же результат, текст задан явно.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $Path,
    [string] $Text
)

$ErrorActionPreference = 'Stop'

if (-not $Text) { $Text = $env:LP_DOC_TEXT }
if (-not $Text) {
    Write-Host '  Documentation оставлен как в git (не задан ни -Text, ни LP_DOC_TEXT)'
    return
}
if (-not (Test-Path -LiteralPath $Path)) { throw "нет файла: $Path" }

$body = [IO.File]::ReadAllText($Path, [Text.UTF8Encoding]::new($false))

$objects = ([regex]::Matches($body, '(?m)^OBJECT ')).Count
if ($objects -eq 0) { throw "в файле нет ни одного OBJECT: $Path" }

# Отступ у содержимого блока - шесть пробелов, как в выгрузке C/SIDE. Хвостовые пробелы
# срезаем: проверка оформления считает их ошибкой, а взяться им неоткуда, кроме как из
# аккуратно отформатированной переменной окружения.
$doc = (($Text -split "`r?`n") | ForEach-Object { ('      ' + $_).TrimEnd() }) -join "`r`n"

$rx = [regex]"(?s)(\r\n    BEGIN\r\n    \{\r\n).*?(\r\n    \}\r\n    END\.\r\n)"
$found = @($rx.Matches($body))
if ($found.Count -ne $objects) {
    throw ("блоков Documentation {0}, а объектов {1} - форма выгрузки изменилась, подмена отменена" -f $found.Count, $objects)
}

# Идём с конца: срезы не сдвигают ещё не обработанные совпадения. Собираем срезами, а не
# Replace, - в тексте подстановки могут быть $ и обратные слэши, и группы бы их съели.
for ($i = $found.Count - 1; $i -ge 0; $i--) {
    $m = $found[$i]
    $body = $body.Substring(0, $m.Index) + $m.Groups[1].Value + $doc + $m.Groups[2].Value +
            $body.Substring($m.Index + $m.Length)
}

[IO.File]::WriteAllText($Path, $body, [Text.UTF8Encoding]::new($false))
Write-Host ("  Documentation заменён в {0} объектах: {1}" -f $objects, ($Text -replace "`r?`n", ' / '))
