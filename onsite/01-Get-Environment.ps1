<#
.SYNOPSIS
    Шаг 1. Снимок окружения сервера: платформа, служба, сборки, права. Один экран.
.DESCRIPTION
    Ничего не меняет. Запускать НА САМОМ сервере NAV.
    Результат снять скриншотом целиком.
#>
param([string]$Instance)
. (Join-Path $PSScriptRoot '_Common.ps1')

Head 'ЛП-01 ОКРУЖЕНИЕ'
$ctx = Get-NavContext -Instance $Instance

Line 'Инстанция'      ($ctx.Instance + '   (все: ' + $ctx.AllServices + ')')
Line 'Служба'         ($ctx.ServiceName + '  ' + $(if ($ctx.Running) { 'работает' } else { 'ОСТАНОВЛЕНА' }))
Line 'Учётка службы'  $ctx.Account
Line 'Версия платформы' $ctx.Version
Line 'Корень NAV'     $ctx.Root
Line 'База'           ($ctx.DbServer + ' / ' + $ctx.DbName)
Line 'Компания'       $ctx.Company
Line 'finsql'         $(if ($ctx.Finsql) { $ctx.Finsql } else { 'НЕ НАЙДЕН - объекты придётся импортировать из C/SIDE вручную' })

# Полная трассировка вызовов C/AL: без неё событие оператора не шлётся вовсе и замер
# выходит пустым. Лежит она в CustomSettings.config инстанции - значит правка требует
# ПЕРЕЗАПУСКА инстанции, а не только галки.
$full = $ctx.Settings['EnableFullALFunctionTracing']
if ($null -eq $full) { $full = '(нет в конфиге)' }
Line 'FullALFunctionTracing' $full

$net = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue)
if ($net) { Line '.NET Framework' ($net.Version + '  (release ' + $net.Release + ')') } else { Line '.NET Framework' '(не прочитан)' }
Line 'csc' $(if (Test-Path $ctx.Csc) { $ctx.Csc } else { 'НЕ НАЙДЕН - приёмник придётся собирать не здесь' })

Head 'СБОРКИ TraceEvent'
$dlls = @(Get-TraceEventDlls -Ctx $ctx)
if ($dlls.Count -eq 0) {
    Bad 'Microsoft.Diagnostics.Tracing.TraceEvent.dll не найдена в Add-ins'
    Bad 'без неё приёмник не соберётся и не запустится - нужна из поставки NAV или от чужого профайлера'
} else {
    foreach ($d in $dlls) {
        Line ('  ' + $d.Version) $d.Path
        # Сумма и вердикт по ней. Версии мало: она говорит, ЧТО связывать, но не говорит,
        # тот ли это файл. Сумма отвечает на второй вопрос сразу и без переписки.
        $known = Get-KnownTraceEvent -Sha $d.Sha
        Line '     сумма' ('{0}  {1} байт' -f $d.Sha.ToLower(), $d.Size)
        if ($known) { Line '     опознана' $known }
        else        { Line '     опознана' 'в известных не значится - приёмник соберётся против неё же, на этом сервере' }
        # Версия тут - версия СБОРКИ, по ней и связывается ссылка; в свойствах файла
        # написана своя, у 2.0.77 она вообще трёхчастная.
        #
        # Цепочка: 1.x самодостаточна, 2.x разложена по отдельным файлам. Каталог
        # надстройки обязан быть самодостаточным, поэтому смотрим её здесь - до того,
        # как шаг 3 что-то куда-то положит.
        $need = @(Get-AssemblyNeeds -Path $d.Path)
        if ($need.Count -eq 0) {
            Line '     цепочка' 'нет - самодостаточна'
        } else {
            Line '     цепочка' (($need | ForEach-Object { $_.Name }) -join ', ')
            $miss = @($need | Where-Object { -not $_.Found })
            if ($miss.Count) { Bad ('  рядом не лежат: ' + (($miss | ForEach-Object { $_.Name }) -join ', ')) }
        }
    }
}

Head 'СОСЕДИ В Add-ins'
# Полный список тут не нужен - на снимок он не влезет; важно лишь, что рядом.
if (Test-Path $ctx.AddIns) {
    $all = @(Get-ChildItem $ctx.AddIns -Directory)
    Line 'Всего каталогов' $all.Count
    $mine = @($all | Where-Object { $_.Name -match 'Etw|Profiler|LineProfiler' })
    if ($mine.Count) { foreach ($m in $mine) { Line '  из них наши/чужие' $m.Name } }
    else { Line '  профайлеров рядом' 'нет' }
} else { Bad ('нет каталога ' + $ctx.AddIns) }

Head 'ПРАВА НА ETW'
# Сессию ETW заводит учётка СЛУЖБЫ, а не тот, кто нажал кнопку. Без прав приёмник
# молча не получит ни одного события. Имя группы берём по SID: на русской Windows
# она называется иначе, а SID один и тот же везде.
$acct = $ctx.Account
$short = $acct
$slash = $short.LastIndexOf([char]92)
if ($slash -ge 0) { $short = $short.Substring($slash + 1) }

if ($acct -eq 'LocalSystem' -or $acct -match 'NT AUTHORITY') {
    Line 'Права на ETW' 'учётка системная - прав достаточно'
} else {
    $grp = Get-CimInstance Win32_Group -Filter "LocalAccount=True AND SID='S-1-5-32-559'" -ErrorAction SilentlyContinue
    if ($grp) {
        $members = (net localgroup $grp.Name) 2>$null
        $inGroup = $false
        foreach ($m in $members) { if ($m -and $m.Trim() -like "*$short*") { $inGroup = $true } }
        Line ('Группа ' + $grp.Name) $(if ($inGroup) { 'учётка службы в ней' } else { 'учётки службы НЕТ - проверить, заводится ли сессия ETW' })
    } else { Warn 'группа Performance Log Users не найдена' }
}
Line 'Учётка службы' $acct
Write-Host '  Окончательный ответ даёт шаг 5: если сессия ETW не заводится, замер выйдет пустым.'

Head 'ЧУЖОЙ ПРОФАЙЛЕР'
# На сервере может уже стоять EtwPerformanceProfiler. Его сборку перезаписывать НЕЛЬЗЯ.
$foreign = @()
if (Test-Path $ctx.AddIns) { $foreign = @(Get-ChildItem $ctx.AddIns -Directory | Where-Object { $_.Name -match 'Etw|Profiler' -and $_.Name -ne 'LineProfiler' }) }
if ($foreign.Count) { foreach ($f in $foreign) { Line '  найден' $f.Name } } else { Line '  найден' 'нет' }
Write-Host ''
