<#
.SYNOPSIS
    Шаг 5. Обкатка профайлера без клиента. Печатает вердикт и ТОЛЬКО упавшие проверки.
.DESCRIPTION
    МЕНЯЕТ базу: заводит настройку профайлера, снимает несколько замеров, сохраняет их
    и включает-выключает журнал покрытия. Ничего чужого не трогает.

    Обкатка судит сама: первая строка - "пройдено N из M". Её и снимать скриншотом.

    ГРАБЛИ, ради которых шаг переписан. Галка "Обкатка" взводится прямым UPDATE в SQL,
    а сервер NAV держит строку настройки в своём кэше и правку мимо себя не видит.
    Кодюнит при этом отрабатывает мгновенно и БЕЗ ОШИБКИ: RunSelfTestIfAsked выходит по
    невзведённой галке, файла нет. Раньше шаг объявлял в этом случае "у учётки службы
    нет прав на запись" - диагноз ложный, и разбор уходил в ACL на полдня. Теперь причина
    определяется уликой: если новых сохранённых замеров в базе не появилось, тело обкатки
    не выполнялось вовсе, и дело в кэше, а не в правах. Лечится сбросом кэша, то есть
    перезапуском экземпляра службы: ключ -RestartService делает это сам.

    Если PowerShell на сервере недоступен - тот же результат даёт клиент:
    настройка профайлера -> галка "Обкатка" -> Object Designer -> Codeunit 110200 -> Run.
    Вердикт покажется одним окном сообщения. Через клиент правка идёт ЧЕРЕЗ NAV, поэтому
    кэш обновляется сам и перезапуск не нужен.
.PARAMETER RestartService
    Перезапустить экземпляр службы перед запуском кодюнита. Нужен, если сервер уже
    прочитал настройку и держит её в кэше. На рабочем сервере рвёт сеансы пользователей,
    поэтому по умолчанию выключен.
#>
param([string]$Instance, [string]$Company, [string]$WorkFolder = 'C:\ProgramData\LineProfiler',
      [switch]$RestartService)
. (Join-Path $PSScriptRoot '_Common.ps1')

Head 'ЛП-05 ОБКАТКА'
$ctx = Get-NavContext -Instance $Instance
if (-not $Company) { $Company = $ctx.Company }
if (-not $Company) {
    $c = @(Invoke-NavSql -Ctx $ctx -Sql 'SELECT TOP 1 Name FROM dbo.[Company] ORDER BY Name')
    if ($c.Count) { $Company = $c[0].Name }
}
Line 'Инстанция' $ctx.Instance
Line 'Компания'  $Company

# Рабочий каталог нужен обкатке для выкладки отчёта. Права на запись нужны учётке СЛУЖБЫ.
if (-not (Test-Path $WorkFolder)) { New-Item -ItemType Directory -Path $WorkFolder | Out-Null }
Line 'Каталог'   $WorkFolder

# Настройка профайлера - таблица без привязки к компании, поэтому имя без префикса.
$has = @(Invoke-NavSql -Ctx $ctx -Sql "SELECT COUNT(*) AS n FROM sys.tables WHERE name = 'AL Profiler Setup'")
if ([int]$has[0].n -eq 0) { Bad 'таблицы настройки нет - сперва шаг 4'; return }
$null = Invoke-NavSqlNonQuery -Ctx $ctx -Sql @"
IF NOT EXISTS (SELECT 1 FROM [AL Profiler Setup]) INSERT INTO [AL Profiler Setup] ([Primary Key]) VALUES ('');
UPDATE [AL Profiler Setup] SET [Self Test Enabled] = 1,
  [Work Folder] = CASE WHEN [Work Folder] = '' THEN '$WorkFolder' ELSE [Work Folder] END;
"@
Line 'Обкатка'   'включена'

# Таблица сохранённых замеров привязана к компании, а имя компании NAV перекладывает
# в имя таблицы со своими заменами - поэтому ищем её, а не собираем строкой.
$runsTable = ''
$cand = @(Invoke-NavSql -Ctx $ctx -Sql 'SELECT name FROM sys.tables WHERE name LIKE ''%$AL Profiler Run''')
foreach ($t in $cand) { if ($t.name -like ($Company + '$AL Profiler Run')) { $runsTable = $t.name; break } }
if (-not $runsTable -and $cand.Count -eq 1) { $runsTable = $cand[0].name }

function Get-RunCount {
    if (-not $runsTable) { return -1 }
    $r = @(Invoke-NavSql -Ctx $ctx -Sql ("SELECT COUNT(*) AS n FROM [{0}]" -f $runsTable))
    if ($r.Count) { return [int]$r[0].n }
    return -1
}
$runsBefore = Get-RunCount

$report = Join-Path $WorkFolder 'selftest.tsv'
if (Test-Path $report) { Remove-Item $report -Force }

$admin = Join-Path $ctx.Root 'NavAdminTool.ps1'
if (-not (Test-Path $admin)) { Bad "нет $admin - запустить Codeunit 110200 из C/SIDE вручную"; return }
# Модуль NAV печатает приветствие прямо в консоль - оно съедает полснимка.
Import-Module $admin -ErrorAction Stop *>$null

if ($RestartService) {
    # Сброс кэша: только перезапуск заставит сервер перечитать настройку из базы.
    Restart-Service $ctx.ServiceName -Force
    $deadline = (Get-Date).AddSeconds(180)
    while ((Get-Date) -lt $deadline) {
        $st = ''
        try { $st = (Get-NAVServerInstance -ServerInstance $ctx.Instance).State } catch { $st = '' }
        if ($st -eq 'Running') { break }
        Start-Sleep -Seconds 2
    }
    Line 'Служба' 'перезапущена, кэш настройки сброшен'
}

$sw = [Diagnostics.Stopwatch]::StartNew()
try { Invoke-NAVCodeunit -ServerInstance $ctx.Instance -CodeunitId 110200 -CompanyName $Company -ErrorAction Stop }
catch { Bad "кодюнит упал: $($_.Exception.Message)" }
$sw.Stop()
Line 'Кодюнит' ('отработал за {0:N0} с' -f $sw.Elapsed.TotalSeconds)

$runsAfter = Get-RunCount
$null = Invoke-NavSqlNonQuery -Ctx $ctx -Sql 'UPDATE [AL Profiler Setup] SET [Self Test Enabled] = 0'
Line 'Обкатка'   'выключена обратно'

Head 'ВЕРДИКТ'
if (-not (Test-Path $report)) {
    Bad "нет $report"
    # Улика: обкатка сохраняет замеры. Не появилось ни одного - тело не выполнялось.
    if ($runsBefore -ge 0 -and $runsAfter -eq $runsBefore) {
        Bad 'новых замеров в базе не появилось - тело обкатки не выполнялось вовсе'
        Bad 'причина не в правах: сервер не увидел взведённую галку, строка настройки у него в кэше'
        if ($RestartService) {
            Line 'Что делать' 'служба уже перезапускалась - взведите галку "Обкатка" в клиенте и запустите Codeunit 110200 оттуда'
        } else {
            Line 'Что делать' 'повторить шаг с ключом -RestartService (перезапустит службу и порвёт сеансы) либо взвести галку в клиенте'
        }
    }
    elseif ($runsAfter -gt $runsBefore) {
        Bad ('замеры прошли (добавилось {0}), а файл не появился - дело в записи' -f ($runsAfter - $runsBefore))
        Bad 'проверить права учётки службы на каталог и значение "Рабочий каталог" в настройке'
    }
    else {
        Bad 'причину определить не удалось: таблица замеров не найдена'
        Bad 'смотреть права учётки службы на каталог и кэш настройки (перезапуск службы)'
    }
    return
}
$lines = Get-Content $report -Encoding UTF8
Write-Host ('  ' + $lines[0]) -ForegroundColor Cyan
$failed = @($lines | Where-Object { $_ -like 'FAILED*' })
if ($failed.Count -eq 0) { Good 'упавших проверок нет' }
else { foreach ($f in $failed) { Write-Host ('  ' + $f) -ForegroundColor Red } }

Head 'ЧТО СМОТРЕТЬ ПЕРВЫМ'
# Смещение номера строки - единственный отказ, который выглядит как успех.
$shift = @($lines | Where-Object { $_ -match 'смещение номера строки|line offset' })
if ($shift.Count) { foreach ($s in $shift) { Write-Host ('  ' + $s) } }
$health = @($lines | Where-Object { $_ -match 'здоровье замера|trace health' })
if ($health.Count) { Write-Host ('  ' + $health[0]) }
Write-Host ''
