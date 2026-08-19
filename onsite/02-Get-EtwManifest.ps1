<#
.SYNOPSIS
    Шаг 2. Манифест провайдера ETW: маски и ИМЕНА полей событий. Один экран.
.DESCRIPTION
    Ничего не меняет. Запускать НА САМОМ сервере NAV: манифест регистрирует установщик.
    Это единственный способ сверить полезную нагрузку целевого сервера с той, на которой
    разрабатывали: приёмник читает поля ПО ИМЕНАМ, и расхождение имён видно только здесь.
#>
. (Join-Path $PSScriptRoot '_Common.ps1')

Head 'ЛП-02 МАНИФЕСТ ETW'
$p = $null
try { $p = Get-WinEvent -ListProvider Microsoft-DynamicsNAV-Server -ErrorAction Stop } catch { }
if (-not $p) {
    Bad 'провайдер Microsoft-DynamicsNAV-Server не зарегистрирован на этой машине'
    Bad 'запускать надо НА сервере NAV, а не на клиенте'
    return
}

Line 'Провайдер' $p.Name
Line 'GUID'      $p.Id
Line 'Событий'   $p.Events.Count

Head 'МАСКИ (keywords)'
foreach ($k in $p.Keywords) {
    if ($k.Value -gt 0 -and $k.Value -lt 0x1000000000000000) {
        Write-Host ('  0x{0:X4}  {1}' -f $k.Value, $k.Name)
    }
}

Head 'СОБЫТИЯ, КОТОРЫЕ ЧИТАЕТ ПРИЁМНИК'
# Отбираем не по номерам (они могут отличаться), а по СМЫСЛУ - по полям нагрузки.
# Одинаковые наборы полей сводим в одну строку: десяток парных событий SQL отличается
# только номером, и на снимок их полный перечень не влезет.
$want = 'functionName|sqlStatement|lineNumber'
$groups = @{}
foreach ($e in ($p.Events | Sort-Object Id)) {
    $t = $e.Template
    if (-not $t) { continue }
    if ($t -notmatch $want) { continue }
    $names = @()
    foreach ($m in [regex]::Matches($t, 'name\s*=\s*"([^"]+)"')) { $names += $m.Groups[1].Value }
    $sig = ($names -join ',')
    if (-not $groups.ContainsKey($sig)) { $groups[$sig] = @() }
    $groups[$sig] += ('{0}v{1}' -f $e.Id, $e.Version)
}
if ($groups.Count -eq 0) { Bad 'событий с нужными полями нет - манифест другой, нужен разбор' }
foreach ($sig in $groups.Keys) {
    Write-Host ('  ' + $sig) -ForegroundColor Green
    Write-Host ('       события: ' + (($groups[$sig]) -join ' '))
}

Head 'ЧТО ЖДЁТ ПРИЁМНИК'
Write-Host '  оператор : objectType objectId functionName lineNumber statement'
Write-Host '  вход/выход: objectType objectId functionName'
Write-Host '  SQL      : sessionId sqlStatement'
Write-Host '  маски    : SqlTracing 0x2 + ALFunctionCallTracing 0x8  (итого 0xA)'
Write-Host ''
Write-Host '  Расхождение ИМЁН - единственное, что придётся править в приёмнике.'
Write-Host '  Расхождение НОМЕРОВ событий и версий шаблонов приёмнику безразлично.'
Write-Host ''
