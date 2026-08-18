#requires -Version 5.1
<#
.SYNOPSIS
    Регресс агрегатора: прогон сохранённых файлов .etl через приёмник профайлера.

.DESCRIPTION
    Скармливает приёмнику ранее собранные трассировки и проверяет инварианты, которые
    на живом прогоне проверить нечем: склейку кадров, парность операций SQL, сходимость
    времени. Живая база и служба NST для этого не нужны — только сборка и файлы.

    ВАЖНО: запускать под Windows PowerShell 5.1 (powershell.exe), а НЕ под pwsh 7.
    Под .NET Core разбор файла падает на P/Invoke:
        Unable to find an entry point named 'ZeroMemory' in DLL 'kernel32.dll'
    NST работает на .NET Framework, поэтому проверять надо именно на нём.

    Файл этот в кодировке UTF-8 С СИГНАТУРОЙ: без неё WinPS 5.1 читает его как ANSI
    и ломается на кириллице в кавычках.

.PARAMETER Path
    Каталог с подпапками прогонов (по умолчанию <корень репозитория>/out) либо
    конкретный файл .etl.

.PARAMETER Dll
    Сборка приёмника. По умолчанию берётся выложенная в Add-ins службы NST, а если её
    там нет — собранная в <корень репозитория>/bin.

.EXAMPLE
    powershell -File scripts\Test-EtlReplay.ps1
#>
[CmdletBinding()]
param(
    [string]$Path,
    [string]$Dll
)

$ErrorActionPreference = 'Stop'

$taskDir = Split-Path $PSScriptRoot -Parent
if (-not $Path) { $Path = Join-Path $taskDir 'out' }

if (-not $Dll) {
    $deployed = 'C:\Program Files\Microsoft Dynamics NAV\110\Service\Add-ins\LineProfiler\AlLineProfiler.dll'
    $built    = Join-Path $taskDir 'bin\AlLineProfiler.dll'
    if (Test-Path $deployed) { $Dll = $deployed } elseif (Test-Path $built) { $Dll = $built }
}
if (-not $Dll -or -not (Test-Path $Dll)) {
    throw "Сборка приёмника не найдена. Соберите её через Deploy-LineProfiler.ps1 или укажите -Dll."
}

# Ссылку на TraceEvent берём из той же папки, где лежит сборка: версии обязаны совпадать,
# сборка со строгим именем связывается по ТОЧНОЙ версии.
$traceRef = Join-Path (Split-Path $Dll -Parent) 'Microsoft.Diagnostics.Tracing.TraceEvent.dll'
if (-not (Test-Path $traceRef)) {
    $traceRef = Get-ChildItem 'C:\Program Files\Microsoft Dynamics NAV\110\Service\Add-ins' `
        -Recurse -Filter 'Microsoft.Diagnostics.Tracing.TraceEvent.dll' -ErrorAction SilentlyContinue |
        Sort-Object { [version]$_.VersionInfo.FileVersion } -Descending |
        Select-Object -First 1 -ExpandProperty FullName
}
if (-not $traceRef) { throw "Microsoft.Diagnostics.Tracing.TraceEvent.dll не найдена." }

Add-Type -Path $traceRef | Out-Null
Add-Type -Path $Dll | Out-Null

if (Test-Path $Path -PathType Leaf) {
    $files = @(Get-Item $Path)
} else {
    $files = @(Get-ChildItem $Path -Recurse -Filter '*.etl' -ErrorAction SilentlyContinue)
}
if (-not $files) { throw "Файлов .etl не найдено: $Path" }

Write-Output ("Сборка : {0}" -f $Dll)
Write-Output ("Файлов : {0}" -f $files.Count)
Write-Output ''

$fail = 0
$checks = 0

function Assert-Invariant([string]$Name, [bool]$Ok, [string]$Detail) {
    $script:checks++
    if (-not $Ok) { $script:fail++ }
    $mark = if ($Ok) { 'ПРОЙДЕНО' } else { 'ПРОВАЛ  ' }
    Write-Output ("    {0} {1}{2}" -f $mark, $Name, $(if ($Detail) { " - $Detail" } else { '' }))
}

foreach ($f in $files) {
    $r = New-Object LineProfiler.AlLineProfilerReceiver
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $rows = $r.ProcessEtlFile($f.FullName)
    $sw.Stop()

    Write-Output ("{0}  ({1:N1} МБ)" -f $f.FullName, ($f.Length / 1MB))
    Write-Output ("    строк {0}, событий {1}, операторов {2}, SQL {3}, за {4} мс" -f `
        $rows, $r.EventsSeen, $r.StatementEvents, $r.SqlEvents, $sw.ElapsedMilliseconds)

    # Инварианты, которые на живом прогоне проверить нечем.
    Assert-Invariant 'результат не пуст' ($rows -gt 0) "строк $rows"
    Assert-Invariant 'потерь буфера нет' ($r.LostEvents -eq 0) "потеряно $($r.LostEvents)"
    # Один-два незакрытых кадра - это запись, оборванная посреди вызова, и для снятой
    # трассы это норма. Систематическая поломка стека выглядит иначе: кадры не снимаются
    # вовсе, и счёт идёт на десятки.
    Assert-Invariant 'стек снимается' ($r.UnclosedFrames -le 8) "осталось открытых $($r.UnclosedFrames)"
    Assert-Invariant 'операции SQL парны' ($r.UnpairedSql -eq 0) "без начала $($r.UnpairedSql)"
    Assert-Invariant 'имена полей разобраны' ($r.UnnamedEvents -eq 0) "безымянных $($r.UnnamedEvents)"

    # Время строки не может быть отрицательным, а время SQL - частью себя не большей.
    $badTime = 0
    $badSql  = 0
    $withText = 0
    for ($i = 0; $i -lt $r.RowCount; $i++) {
        [void]$r.SelectRow($i)
        if ($r.CurrentSelfMs -lt 0 -or $r.CurrentTotalMs -lt 0) { $badTime++ }
        if ($r.CurrentSqlMs -gt $r.CurrentSelfMs + 0.5) { $badSql++ }
        if ($r.CurrentLineText) { $withText++ }
    }
    Assert-Invariant 'время неотрицательно' ($badTime -eq 0) "строк с минусом $badTime"
    Assert-Invariant 'время SQL не больше времени строки' ($badSql -eq 0) "нарушений $badSql"
    Assert-Invariant 'текст оператора есть у каждой строки' ($withText -eq $rows) "$withText из $rows"

    # Тексты запросов: выборка без отбора обязана вернуть ровно то, что собрано.
    $q = $r.SelectQueries(-1, -1, -1, '')
    Assert-Invariant 'выборка запросов сходится со счётчиком' ($q -eq $r.SqlRowsCollected) `
        "$q против $($r.SqlRowsCollected)"

    Write-Output ''
}

Write-Output ("ИТОГ: пройдено {0} из {1}" -f ($checks - $fail), $checks)
if ($fail -gt 0) { exit 1 }
