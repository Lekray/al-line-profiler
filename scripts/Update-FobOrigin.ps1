#requires -Version 7
<#
.SYNOPSIS
    Обновляет fob-origin.txt — объявление, из какого текста снят LineProfiler.fob.

.DESCRIPTION
    Запускается СРАЗУ ПОСЛЕ того, как .fob перевыгружен из C/SIDE с базы, в которую
    импортирован текущий LineProfiler.txt. Другого повода нет: манифест — это не
    отчёт о состоянии файлов, а показание о происхождении контейнера.

    Собрать .fob из .txt скриптом нельзя, это делает C/SIDE, поэтому доказать
    происхождение офлайн нечем — только объявить и потом сверять. Зато объявление
    проверяемое: любая последующая правка текста ломает хэш, и Test-ObjectFile.ps1
    говорит, что контейнер отстал.

    Состав объектов сверяется до записи: если оглавление .fob и текст разошлись по
    объектам, именам, датам или спискам версий, манифест не пишется вовсе — сначала
    разобраться, что за контейнер лежит рядом.

    Ключ -Force нужен ровно в опасном случае: текст изменился с прошлого объявления.
    Без него скрипт откажется — чтобы «обновить манифест» нельзя было сделать вместо
    перевыгрузки, случайно превратив проверку в ложь.

.PARAMETER TaskFile
    Текстовый экспорт. По умолчанию LineProfiler.txt в корне репозитория.

.PARAMETER Platform
    Версия платформы, которой снят контейнер. По умолчанию — то, что уже записано
    в манифесте.

.PARAMETER Force
    Разрешить перезапись, когда текст изменился с прошлого объявления.

.EXAMPLE
    pwsh scripts/Update-FobOrigin.ps1 -Force
    После перевыгрузки .fob: объявить новое происхождение.
#>
[CmdletBinding()]
param(
    [string] $TaskFile,
    [string] $Platform,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Lib-CSideFob.ps1')

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $TaskFile) { $TaskFile = Join-Path $repoRoot 'LineProfiler.txt' }
if (-not (Test-Path -LiteralPath $TaskFile -PathType Leaf)) { throw "Нет файла: $TaskFile" }

$src        = (Resolve-Path -LiteralPath $TaskFile).Path
$fobPath    = [System.IO.Path]::ChangeExtension($src, '.fob')
$originPath = Join-Path (Split-Path -Parent $src) 'fob-origin.txt'
if (-not (Test-Path -LiteralPath $fobPath -PathType Leaf)) { throw "Нет контейнера: $fobPath" }

# --- состав должен совпасть, иначе объявлять нечего -------------------------
$fobDir = @(Get-CSideFobDirectory -Path $fobPath)
if ($fobDir.Count -eq 0) { throw "В $fobPath не нашлось оглавления объектов." }
$txtDir = @(Get-CSideTextDirectory -Path $src)
$diff   = @(Compare-CSideDirectory -Fob $fobDir -Text $txtDir)
if ($diff.Count -gt 0) {
    Write-Host 'Состав контейнера расходится с текстом — манифест не записан:' -ForegroundColor Red
    foreach ($d in $diff) { Write-Host ('  ' + $d) -ForegroundColor Red }
    exit 1
}

# --- прошлое объявление ------------------------------------------------------
$prev = @{}
$head = New-Object System.Collections.Generic.List[string]
if (Test-Path -LiteralPath $originPath -PathType Leaf) {
    foreach ($line in [System.IO.File]::ReadAllLines($originPath, [System.Text.Encoding]::UTF8)) {
        if ($line.TrimStart().StartsWith('#') -or -not $line.Trim()) { [void]$head.Add($line); continue }
        $m = [regex]::Match($line, '^\s*([a-z]+)\s+(\S+)\s*$')
        if ($m.Success) { $prev[$m.Groups[1].Value] = $m.Groups[2].Value }
    }
}

$shaTxt = Get-CSideSha256 $src
$shaFob = Get-CSideSha256 $fobPath
if (-not $Platform) {
    if ($prev.ContainsKey('platform')) { $Platform = $prev['platform'] } else { $Platform = 'неизвестна' }
}

if ($prev.ContainsKey('txt') -and $prev['txt'] -ne $shaTxt -and -not $Force) {
    Write-Host 'Текст изменился с прошлого объявления.' -ForegroundColor Yellow
    Write-Host ('  было:   {0}' -f $prev['txt'])
    Write-Host ('  сейчас: {0}' -f $shaTxt)
    Write-Host ''
    Write-Host 'Если .fob ТОЛЬКО ЧТО перевыгружен с базы, куда импортирован этот текст —' -ForegroundColor Yellow
    Write-Host 'повторите с ключом -Force. Если нет — сначала перевыгрузите контейнер:' -ForegroundColor Yellow
    Write-Host 'обновление манифеста вместо перевыгрузки делает проверку ложной.' -ForegroundColor Yellow
    exit 1
}

# Шапка с пояснением сохраняется: без неё манифест выглядит набором хэшей, и
# соблазн «обновить, чтобы позеленело» возвращается на следующий же раз.
if ($head.Count -eq 0) {
    [void]$head.Add('# Происхождение LineProfiler.fob — с какого текста он снят.')
    [void]$head.Add('#')
    [void]$head.Add('# txt  — SHA-256 текстового экспорта на момент выгрузки .fob из C/SIDE.')
    [void]$head.Add('# fob  — SHA-256 самого контейнера.')
    [void]$head.Add('#')
    [void]$head.Add('# Обновляется scripts/Update-FobOrigin.ps1 СРАЗУ ПОСЛЕ перевыгрузки, и только тогда.')
    [void]$head.Add('')
}
while ($head.Count -gt 0 -and -not $head[$head.Count - 1].Trim()) { $head.RemoveAt($head.Count - 1) }

$out = New-Object System.Collections.Generic.List[string]
foreach ($h in $head) { [void]$out.Add($h) }
[void]$out.Add('')
[void]$out.Add(('txt      {0}' -f $shaTxt))
[void]$out.Add(('fob      {0}' -f $shaFob))
[void]$out.Add(('platform {0}' -f $Platform))

[System.IO.File]::WriteAllText($originPath, (($out -join "`r`n") + "`r`n"),
                               (New-Object System.Text.UTF8Encoding($false)))

Write-Host 'Объявлено:'
Write-Host ('  txt      {0}  ({1})' -f $shaTxt, [System.IO.Path]::GetFileName($src))
Write-Host ('  fob      {0}  ({1})' -f $shaFob, [System.IO.Path]::GetFileName($fobPath))
Write-Host ('  platform {0}' -f $Platform)
Write-Host ('Объектов сверено: {0}' -f $fobDir.Count)
Write-Host ''
Write-Host ('Не забудьте про README: там опубликован SHA-256 контейнера — {0}' -f $shaFob) -ForegroundColor Yellow
