<#
.SYNOPSIS
    Шаг 4. Импорт и компиляция объектов профайлера, с проверкой по базе.
.DESCRIPTION
    МЕНЯЕТ базу: заводит 16 объектов диапазона 110200-110207.
    Импортируется файл в кодировке cp866 - другие кодировки ломают кириллицу в C/SIDE.

    Грабли, ради которых написан скрипт: импорт СБРАСЫВАЕТ флаг Compiled. Если
    скомпилировать не все объекты, на первом же обращении к недокомпилированной таблице
    сервер отвечает "Объект метаданных Table N не найден" и в лог ничего не пишет.
    Поэтому компилируются ВСЕ, а результат сверяется с dbo.[Object].
#>
param([string]$Instance, [string]$Company)
. (Join-Path $PSScriptRoot '_Common.ps1')

Head 'ЛП-04 ОБЪЕКТЫ'
$ctx = Get-NavContext -Instance $Instance
$txt = Join-Path $PSScriptRoot 'objects\LineProfiler.cp866.txt'
if (-not (Test-Path $txt)) { Bad "нет файла объектов: $txt"; return }
if (-not $ctx.Finsql) { Bad 'finsql не найден - импортировать из C/SIDE вручную (см. README)'; return }

Line 'База'   ($ctx.DbServer + ' / ' + $ctx.DbName)
Line 'Файл'   $txt

$log = Join-Path $env:TEMP 'LineProfiler-import.log'
if (Test-Path $log) { Remove-Item $log -Force }
$common = "ServerName=$($ctx.DbServer),Database=$($ctx.DbName),LogFile=$log"

& $ctx.Finsql "Command=ImportObjects,File=$txt,$common,SynchronizeSchemaChanges=No" | Out-Null
if (Test-Path $log) { Bad 'импорт с замечаниями:'; Get-Content $log | Select-Object -First 12 | ForEach-Object { Write-Host ('    ' + $_) } }
else { Good 'импорт прошёл' }

# Компилируем ВЕСЬ диапазон разом: порядок C/SIDE разрешает сам.
$filter = 'Type=1|3|5|8;ID=110200..110207'
if (Test-Path $log) { Remove-Item $log -Force }
& $ctx.Finsql "Command=CompileObjects,Filter=$filter,$common,SynchronizeSchemaChanges=Force" | Out-Null
if (Test-Path $log) { Bad 'компиляция с замечаниями:'; Get-Content $log | Select-Object -First 12 | ForEach-Object { Write-Host ('    ' + $_) } }

Head 'ПРОВЕРКА ПО БАЗЕ'
$rows = @(Invoke-NavSql -Ctx $ctx -Sql @"
SELECT Type, ID, Name, Compiled FROM dbo.[Object]
WHERE ID BETWEEN 110200 AND 110207 AND Type IN (1,3,5,8) ORDER BY Type, ID
"@)
if ($rows.Count -eq 0) { Bad 'в базе нет ни одного нашего объекта - импорт не прошёл'; return }
$names = @{ 1 = 'Table'; 3 = 'Report'; 5 = 'Codeunit'; 8 = 'Page' }
$bad = 0
foreach ($r in $rows) {
    $ok = ($r.Compiled -eq 1)
    if (-not $ok) { $bad++ }
    Write-Host ('  {0,-9} {1}  {2,-28} {3}' -f $names[[int]$r.Type], $r.ID, $r.Name, $(if ($ok) { 'скомпилирован' } else { 'НЕ СКОМПИЛИРОВАН' }))
}
Write-Host ''
Line 'Всего объектов' $rows.Count
Line 'Не скомпилировано' $bad
if ($bad -eq 0 -and $rows.Count -eq 16) { Good 'все 16 объектов на месте' } else { Bad 'разобраться до запуска замера' }
Write-Host ''
