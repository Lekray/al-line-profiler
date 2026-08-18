<#
.SYNOPSIS
    Шаг 3. Сборка приёмника ETW НА МЕСТЕ и выкладка в Add-ins службы.
.DESCRIPTION
    МЕНЯЕТ систему: кладёт свою сборку в Add-ins и ПЕРЕЗАПУСКАЕТ службу NAV.

    Собирать надо именно на месте. Ссылка на TraceEvent строгоименная, то есть
    связывается только с ТОЧНОЙ версией: сборка против 1.0.39 против 1.0.11 не
    заработает. Сборка против той библиотеки, что лежит на ЭТОМ сервере, снимает
    вопрос совместимости по построению. Исходник этого не замечает - он собирается
    без правок против 1.0.11, 1.0.26, 1.0.39 и 2.0.64.

    Если компилятора на сервере нет, берётся готовая сборка из пакета - тот её
    вариант, что отвечает версии TraceEvent на этом сервере.

    Саму TraceEvent пакет НЕ везёт: её версии до 2.0.30 лицензия Microsoft
    распространять не разрешает, да это и не нужно - библиотека уже здесь.
    Установка копирует её РЯДОМ с приёмником: NAV не просматривает соседние
    папки надстроек, каждая обязана быть самодостаточной.

    Чужую сборку (EtwPerformanceProfiler) НЕ трогает: своя ложится в свою подпапку.
    Прежняя своя сборка сохраняется рядом и возвращается при неудаче.
#>
param([string]$Instance, [switch]$NoRestart)
. (Join-Path $PSScriptRoot '_Common.ps1')

Head 'ЛП-03 СБОРКА ПРИЁМНИКА'
$ctx = Get-NavContext -Instance $Instance

$src      = Join-Path $PSScriptRoot 'receiver\AlLineProfiler.cs'
$variants = Join-Path $PSScriptRoot 'receiver\variants'

$outDir = Join-Path $env:TEMP 'LineProfiler-build'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$dll = Join-Path $outDir 'AlLineProfiler.dll'

$dlls = @(Get-TraceEventDlls -Ctx $ctx)
if ($dlls.Count -eq 0) { Bad 'TraceEvent.dll на сервере не найдена - работать не с чем (см. шаг 1)'; return }
foreach ($d in $dlls) { Line 'TraceEvent' ('{0,-10} {1}' -f $d.Version, $d.Path) }
$te = $dlls[0]

# Сравниваем три первые части: FileVersionInfo даёт «1.0.39.0», в имени файла «1.0.39».
function Test-SameTeVersion([string]$A, [string]$B) {
    $x = [version]$A; $y = [version]$B
    return ($x.Major -eq $y.Major -and $x.Minor -eq $y.Minor -and $x.Build -eq $y.Build)
}

# Служба отпускает файл не в момент перехода в Stopped, а когда процесс НА САМОМ ДЕЛЕ
# завершился. Ждём освобождения самого файла, а не статуса службы: иначе копирование
# падает с 'used by another process', и выкладка тихо не происходит.
function Wait-FileFree([string]$Path, [int]$Tries = 40) {
    if (-not (Test-Path $Path)) { return $true }
    for ($i = 0; $i -lt $Tries; $i++) {
        try { $fs = [IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $fs.Close(); return $true }
        catch { Start-Sleep -Milliseconds 1500 }
    }
    return $false
}

$built = $false
if ((Test-Path $ctx.Csc) -and (Test-Path $src)) {
    & $ctx.Csc /nologo /target:library /optimize+ /out:"$dll" /reference:"$($te.Path)" "$src" 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        $built = $true
        Line 'Собрано' ("$((Get-Item $dll).Length) байт против $($te.Version)")
    } else {
        Line 'Сборка' "csc вернул $LASTEXITCODE - беру готовую из пакета"
    }
} elseif (-not (Test-Path $ctx.Csc)) {
    Line 'Компилятор' "нет ($($ctx.Csc)) - беру готовую из пакета"
} else {
    Line 'Исходник' "нет ($src) - беру готовую из пакета"
}

if (-not $built) {
    $pick = $null
    if (Test-Path $variants) {
        foreach ($f in (Get-ChildItem $variants -Filter 'AlLineProfiler-TraceEvent-*.dll')) {
            $v = $f.BaseName -replace '^AlLineProfiler-TraceEvent-', ''
            if (Test-SameTeVersion $v $te.Version) { $pick = $f; break }
        }
    }
    if (-not $pick) {
        Bad "готовой сборки под TraceEvent $($te.Version) в пакете нет"
        if (Test-Path $variants) {
            Line 'Есть в пакете' ((Get-ChildItem $variants -Filter '*.dll' | ForEach-Object { $_.BaseName -replace '^AlLineProfiler-TraceEvent-', '' }) -join ', ')
        }
        Line 'Что делать' 'собрать приёмник из src\AlLineProfiler.cs против этой версии - правок в исходнике не нужно'
        return
    }
    Copy-Item $pick.FullName $dll -Force
    Line 'Готовая сборка' ("$($pick.Name), $((Get-Item $dll).Length) байт")
}

$targetDir = Join-Path $ctx.AddIns 'LineProfiler'
$target    = Join-Path $targetDir 'AlLineProfiler.dll'
$traceDst  = Join-Path $targetDir 'Microsoft.Diagnostics.Tracing.TraceEvent.dll'
if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir | Out-Null }
$backup = Join-Path $outDir 'AlLineProfiler.rollback.dll'
$had = Test-Path $target
if ($had) { Copy-Item $target $backup -Force }

# Запущенная служба держит файл: остановить -> копировать -> запустить.
$wasRunning = $ctx.Running
if ($wasRunning -and -not $NoRestart) {
    Stop-Service $ctx.ServiceName -Force
    Line 'Служба' 'остановлена'
}
$copied = $false
if (-not (Wait-FileFree $target)) {
    Bad 'прежняя сборка осталась занятой - копирование пропущено, ничего не изменено'
} else {
    try { Copy-Item $dll $target -Force; $copied = $true; Line 'Выложено' $target }
    catch { Bad "копирование не удалось: $($_.Exception.Message)" }
}

# TraceEvent - в ту же папку, с этого же сервера. Без неё приёмник не поднимется:
# соседнюю папку надстройки платформа не просматривает.
if ($copied) {
    if ([IO.Path]::GetFullPath($te.Path) -eq [IO.Path]::GetFullPath($traceDst)) {
        Line 'TraceEvent' 'уже на месте'
    } else {
        if (-not (Wait-FileFree $traceDst)) { Bad 'прежняя TraceEvent осталась занятой'; $copied = $false }
        else {
        try { Copy-Item $te.Path $traceDst -Force; Line 'TraceEvent' "положена рядом: $($te.Version)" }
        catch { Bad "TraceEvent не скопировалась: $($_.Exception.Message)"; $copied = $false } }
    }
}

# Кэш сборок C/SIDE привязан к имени папки и версии; версию мы держим неизменной,
# иначе объявление DotNet в C/AL пришлось бы править на каждую пересборку.
$cache = Join-Path $env:TEMP 'Microsoft Dynamics NAV\Add-Ins'
if (Test-Path $cache) {
    Get-ChildItem $cache -Recurse -Directory -Filter 'LineProfiler*' -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    Line 'Кэш C/SIDE' 'очищен'
}

if ($wasRunning -and -not $NoRestart) {
    try { Start-Service $ctx.ServiceName; Line 'Служба' 'запущена' }
    catch {
        Bad 'служба не поднялась - возвращаю прежнюю сборку'
        if ($had) { Copy-Item $backup $target -Force }
        Start-Service $ctx.ServiceName
    }
}
if ($copied) { Good 'приёмник на месте' }
Write-Host ''
