<#
.SYNOPSIS
    Шаг 5. Обкатка профайлера без клиента. Печатает вердикт и ТОЛЬКО упавшие проверки.
.DESCRIPTION
    МЕНЯЕТ базу: заводит настройку профайлера, снимает несколько замеров, сохраняет их
    и включает-выключает журнал покрытия. Ничего чужого не трогает.

    Обкатка судит сама: первая строка - "пройдено N из M". Её и снимать скриншотом.

    Если PowerShell на сервере недоступен - тот же результат даёт клиент:
    настройка профайлера -> галка "Обкатка" -> Object Designer -> Codeunit 110200 -> Run.
    Вердикт покажется одним окном сообщения.
#>
param([string]$Instance, [string]$Company, [string]$WorkFolder = 'C:\ProgramData\LineProfiler')
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

$report = Join-Path $WorkFolder 'selftest.tsv'
if (Test-Path $report) { Remove-Item $report -Force }

$admin = Join-Path $ctx.Root 'NavAdminTool.ps1'
if (-not (Test-Path $admin)) { Bad "нет $admin - запустить Codeunit 110200 из C/SIDE вручную"; return }
# Модуль NAV печатает приветствие прямо в консоль - оно съедает полснимка.
Import-Module $admin -ErrorAction Stop *>$null
try { Invoke-NAVCodeunit -ServerInstance $ctx.Instance -CodeunitId 110200 -CompanyName $Company -ErrorAction Stop }
catch { Bad "кодюнит упал: $($_.Exception.Message)" }

$null = Invoke-NavSqlNonQuery -Ctx $ctx -Sql 'UPDATE [AL Profiler Setup] SET [Self Test Enabled] = 0'
Line 'Обкатка'   'выключена обратно'

Head 'ВЕРДИКТ'
if (-not (Test-Path $report)) {
    Bad "нет $report"
    Bad 'обычная причина - у учётки службы нет прав на запись в каталог'
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
