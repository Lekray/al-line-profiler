#requires -Version 7
<#
.SYNOPSIS
    Собирает ОДИН объект C/AL, который несёт в себе сборку приёмника и ставит её.

.DESCRIPTION
    Нужен там, где нет ни доступа к диску сервера, ни возможности гонять письма кругами.
    Сборка приёмника (29 КБ) едет ВНУТРИ объекта текстом, поэтому отдельного вложения,
    отдельной рассылки частей и ручного копирования не требуется вовсе.

    TraceEvent.dll почтой НЕ едет: она весит 2 МБ, это шесть десятков писем. Её на
    сервере копирует сама служба - из той папки Add-ins, где найдёт, в нашу. Чужую
    папку объект не трогает.

    На выходе - один .txt в кодировке ASCII без BOM: тело письма сохраняется как есть
    и импортируется в C/SIDE.

.EXAMPLE
    pwsh scripts/New-InstallerObject.ps1
#>
[CmdletBinding()]
param([string]$OutDir)

$ErrorActionPreference = 'Stop'

$taskDir = Split-Path $PSScriptRoot -Parent
$tpl     = Join-Path $taskDir 'onsite\Installer.template.txt'
# Берём ОПУБЛИКОВАННУЮ сборку из dist, а не свежую из bin: именно по dist согласована
# установка и посчитаны контрольные суммы, и в почтовый объект должно попасть ровно то,
# что названо в согласовании. Так же поступают оба сборщика пакетов.
$dll     = Join-Path $taskDir 'dist\AlLineProfiler.dll'
if (-not $OutDir) { $OutDir = Join-Path $taskDir 'out' }

function Step([string]$m) { Write-Host "  $m" }
Write-Host 'Установщик приёмника одним объектом' -ForegroundColor Cyan

foreach ($p in @($tpl, $dll)) { if (-not (Test-Path $p)) { throw "не найдено: $p" } }
$bin = Join-Path $taskDir 'bin\AlLineProfiler.dll'
if ((Test-Path $bin) -and ((Get-FileHash $bin).Hash -ne (Get-FileHash $dll).Hash)) {
    throw "bin\AlLineProfiler.dll отличается от dist\AlLineProfiler.dll. В объект идёт dist. Обновите dist и пересчитайте dist\SHA256SUMS.txt либо удалите bin."
}

$bytes = [IO.File]::ReadAllBytes($dll)
$b64   = [Convert]::ToBase64String($bytes)
Step ("сборка $($bytes.Length) байт -> $($b64.Length) знаков base64")

# Нагрузка разложена по нескольким процедурам: у C/AL есть предел на размер одной
# функции, и упереться в него на чужой платформе было бы обиднее всего.
$perLine = 64
$perProc = 100
$lines = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $b64.Length; $i += $perLine) {
    $lines.Add($b64.Substring($i, [Math]::Min($perLine, $b64.Length - $i)))
}
$procCount = [Math]::Ceiling($lines.Count / $perProc)
Step ("строк $($lines.Count), процедур $procCount")

$sb = New-Object System.Text.StringBuilder
$callList = @()
for ($p = 0; $p -lt $procCount; $p++) {
    $id = 6 + $p
    $name = 'Payload{0:d2}' -f ($p + 1)
    $callList += $name
    [void]$sb.AppendLine("    LOCAL PROCEDURE $name@$id(VAR B@1000000000 : DotNet ""'mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'.System.Text.StringBuilder"");")
    [void]$sb.AppendLine('    BEGIN')
    $take = [Math]::Min($perProc, $lines.Count - $p * $perProc)
    for ($k = 0; $k -lt $take; $k++) {
        [void]$sb.AppendLine("      B.Append('" + $lines[$p * $perProc + $k] + "');")
    }
    [void]$sb.AppendLine('    END;')
    [void]$sb.AppendLine('')
}

$head = New-Object System.Text.StringBuilder
[void]$head.AppendLine('    LOCAL PROCEDURE AppendPayload@5(VAR B@1000000000 : DotNet "''mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089''.System.Text.StringBuilder");')
[void]$head.AppendLine('    BEGIN')
[void]$head.AppendLine('      // The receiver assembly itself, carried as text so that it can travel in a')
[void]$head.AppendLine('      // mail body. Split across several procedures because C/AL limits how large a')
[void]$head.AppendLine('      // single function may be.')
foreach ($c in $callList) { [void]$head.AppendLine("      $c(B);") }
[void]$head.AppendLine('    END;')
[void]$head.AppendLine('')

$obj = [IO.File]::ReadAllText($tpl).Replace('@@PAYLOAD@@', $head.ToString() + $sb.ToString())
$obj = $obj -replace 'Date=\d\d\.\d\d\.\d\d;', ('Date=' + (Get-Date -Format 'dd.MM.yy') + ';')
$obj = $obj -replace 'DD \d\d\.\d\d\.\d{4} LineProfiler', ('DD ' + (Get-Date -Format 'dd.MM.yyyy') + ' LineProfiler')

$bad = $obj.ToCharArray() | Where-Object { [int]$_ -gt 127 }
if ($bad) { throw "в объекте $($bad.Count) не-ASCII знаков - он не переживёт почту" }

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$out = Join-Path $OutDir 'LineProfiler-installer.txt'
# Без BOM: C/SIDE спотыкается о невидимые байты перед OBJECT.
[IO.File]::WriteAllText($out, $obj, [Text.UTF8Encoding]::new($false))

Write-Host ''
Write-Host ("ИТОГ: $out  ({0:N0} байт)" -f (Get-Item $out).Length) -ForegroundColor Green
