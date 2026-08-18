<#
.SYNOPSIS
    Шаг 6. Права SQL для съёма ЗНАЧЕНИЙ параметров запросов. Необязательный.
.DESCRIPTION
    Ничего не меняет. Значения параметров платформа не отдаёт - они снимаются
    сеансом расширенных событий сервера SQL. Это отдельная возможность: без неё
    профайлер работает целиком, в скрипте для SSMS вместо значений стоит NULL.

    Нужны учётке СЛУЖБЫ NAV: ALTER ANY EVENT SESSION и VIEW SERVER STATE.
#>
param([string]$Instance)
. (Join-Path $PSScriptRoot '_Common.ps1')

Head 'ЛП-06 ПРАВА SQL'
$ctx = Get-NavContext -Instance $Instance
Line 'Сервер'  $ctx.DbServer
Line 'База'    $ctx.DbName
Line 'Учётка службы' $ctx.Account

$v = @(Invoke-NavSql -Ctx $ctx -Sql "SELECT SERVERPROPERTY('ProductVersion') AS v, SERVERPROPERTY('Edition') AS e")
Line 'Версия SQL' ($v[0].v.ToString() + '   ' + $v[0].e)

$acct = $ctx.Account
if ($acct -eq 'LocalSystem') { $acct = 'NT AUTHORITY\SYSTEM' }
$sql = @"
SELECT p.permission_name AS perm
FROM sys.server_permissions p
JOIN sys.server_principals s ON s.principal_id = p.grantee_principal_id
WHERE s.name = '$acct' AND p.permission_name IN ('ALTER ANY EVENT SESSION','VIEW SERVER STATE')
"@
$perm = @()
try { $perm = @(Invoke-NavSql -Ctx $ctx -Sql $sql | ForEach-Object { $_.perm }) }
catch { Warn "права не прочитаны: $($_.Exception.Message)" }

foreach ($need in @('ALTER ANY EVENT SESSION','VIEW SERVER STATE')) {
    if ($perm -contains $need) { Good "$need - есть" } else { Warn "$need - НЕТ" }
}

Head 'ЕСЛИ ПРАВ НЕТ'
Write-Host '  Съём значений выключается галкой в настройке профайлера и на остальное не влияет.'
Write-Host '  Если решат выдать - под sysadmin:'
Write-Host ("    GRANT ALTER ANY EVENT SESSION TO [$acct];")
Write-Host ("    GRANT VIEW SERVER STATE TO [$acct];")
Write-Host ''
