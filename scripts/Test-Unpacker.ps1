#Requires -Version 5.1
<#
.SYNOPSIS
    Прогон распаковщика base64 (Codeunit 110207) на стенде, без клиента.

.DESCRIPTION
    Единственная проверка в наборе, которой нужна база: распаковщик - объект C/AL, и
    снаружи NAV он не исполняется никак. Зато исполняется БЕЗ клиента: без GUI объект
    читает файлы из подкаталога СВОЕГО рабочего каталога (TEMPORARYPATH\unpack), а не
    спрашивает их диалогом.

    Каталог именно свой, а не заданный снаружи: объект, который распаковывает что
    угодно, не имеет права знать ни одного чужого пути. Раньше он перебирал четыре
    зашитых пути в поисках файла с именем part01.txt - и это было незаметно, потому
    что с клиентом он файлы СПРАШИВАЕТ.

    Сверяется результат, а не вид: собранный файл сличается по SHA-256 с тем, из
    которого части и сделаны.

      1. две помеченные части по порядку;
      2. те же части, но номера в них ОБРАТНЫ именам файлов - объект обязан
         переставить их сам. base64, склеенный не в том порядке, остаётся валидным
         base64: раскодируется молча и даёт мусор, поэтому случай и заведён;
      3. чистый base64 одним файлом, без меток и шапки - так его отдают чужие
         системы, и распаковщик обязан принимать и такое;
      4. обрезанная часть: объект обязан ОТКАЗАТЬ, назвать причину и не упасть, а
         СЛЕДУЮЩИЙ запуск сразу за ним - пройти. Это главный случай: прежняя версия
         звала ERROR посреди чтения, файл части оставался открытым на службе, и
         следующий запуск вставал на нём намертво;
      5. посторонний текст вместо нагрузки - отказ с причиной;
      6. каталог есть, но пуст - отказ с причиной;
      7. каталога нет вовсе - отказ, и в отчёте сказано, КАКОЙ каталог нужен;
      8. имена файлов произвольные, ничего похожего на partNN.txt - собирается;
      9. имя файла приехало строкой #NAME - объект обязан его взять;
     10. в #NAME путь наружу - имя отвергается целиком, а не чинится;
     11. в #NAME кириллица - то же самое: объект сплошь латинский;
     12. в #NAME расширение спорит с сигнатурой - верх берёт имя (docx - это zip).

    Прогон печатает «пройдено N из M» и сам решает, сошлось ли.

.PARAMETER ServerInstance
    Экземпляр службы NAV.

.PARAMETER CompanyName
    Компания, в которой запускать объект.

.PARAMETER Database
    База стенда. По умолчанию из LP_DATABASE.

.PARAMETER Import
    Перед прогоном импортировать и скомпилировать onsite/Unpacker.template.txt.

.EXAMPLE
    powershell -File scripts\Test-Unpacker.ps1 -CompanyName CRONUS -Import
#>
[CmdletBinding()]
param(
    [string] $ServerInstance = 'DynamicsNAV110',
    [string] $CompanyName,
    [string] $Database = $(if ($env:LP_DATABASE) { $env:LP_DATABASE } else { 'NAV' }),
    [switch] $Import
)

$ErrorActionPreference = 'Stop'
$repo    = Split-Path $PSScriptRoot -Parent
$navRoot = 'C:\Program Files\Microsoft Dynamics NAV\110\Service'
$admin   = Join-Path $navRoot 'NavAdminTool.ps1'
$finsql  = 'C:\Program Files (x86)\Microsoft Dynamics NAV\110\RoleTailored Client\finsql.exe'
$navUsers = 'C:\ProgramData\Microsoft\Microsoft Dynamics NAV\110\Server\' +
            'MicrosoftDynamicsNavServer$' + $ServerInstance + '\users'

$script:Passed = 0
$script:Total  = 0
$script:Fails  = New-Object System.Collections.Generic.List[string]

function Test-Value {
    param([string]$Name, $Expected, $Actual)
    $script:Total++
    if ($Actual -eq $Expected) { $script:Passed++ }
    else { [void]$script:Fails.Add(("FAILED {0}: ожидалось {1}, получено {2}" -f $Name, $Expected, $Actual)) }
}

# TEMPORARYPATH у NAV - не temp учётной записи службы, а профиль ЭКЗЕМПЛЯРА:
# ...\Server\<экземпляр>\users\default\<домен>\<пользователь>\TEMP. Пользователь там
# свой у каждой сессии, поэтому каталог ищется, а не задаётся.
function Get-NavTempDirs {
    $d = @()
    if (Test-Path -LiteralPath $navUsers) {
        $d += @(Get-ChildItem -LiteralPath $navUsers -Recurse -Directory -Filter 'TEMP' -ErrorAction SilentlyContinue |
                ForEach-Object { $_.FullName })
    }
    return $d
}

# Сессий у экземпляра может быть заведено несколько, и какая из них достанется
# запуску - не наше дело. Файлы кладутся во все: объект читает только свой каталог,
# а лишние копии рядом ему не мешают.
function Get-UnpackDirs {
    return @(Get-NavTempDirs | ForEach-Object { Join-Path $_ 'unpack' })
}

# Имена, под которыми объект может оставить результат: 'payload.<тип>', когда имени
# ему не дали, и то, что стояло в #NAME, когда дали.
$script:OutNames = @('payload.*', 'unpacker.log', 'report.zip', 'notes.docx', 'evil.zip', 'my report.zip')

function New-Parts {
    foreach ($d in (Get-UnpackDirs)) {
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
        # Каталог внутри профиля службы, права наследуются - но лишнее разрешение на
        # чтение ничего не стоит, а без прав все случаи сойдутся на «файлов нет»:
        # зелено и бессмысленно.
        & icacls $d /grant '*S-1-5-20:(OI)(CI)R' /T *>$null
    }
}

function Clear-Parts {
    foreach ($d in (Get-UnpackDirs)) {
        if (Test-Path -LiteralPath $d) {
            Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }
}

function Remove-Parts {
    foreach ($d in (Get-UnpackDirs)) {
        if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Clear-Output {
    foreach ($d in (Get-NavTempDirs)) {
        foreach ($n in $script:OutNames) {
            Get-ChildItem (Join-Path $d $n) -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-Output {
    param([string]$Name = 'payload.*')
    $best = $null
    foreach ($d in (Get-NavTempDirs)) {
        foreach ($f in @(Get-ChildItem (Join-Path $d $Name) -ErrorAction SilentlyContinue)) {
            if ($null -eq $best -or $f.LastWriteTime -gt $best.LastWriteTime) { $best = $f }
        }
    }
    return $best
}

function Get-Log {
    foreach ($d in (Get-NavTempDirs)) {
        $p = Join-Path $d 'unpacker.log'
        if (Test-Path -LiteralPath $p) { return (Get-Content -LiteralPath $p -Raw) }
    }
    return ''
}

# Имя файла - произвольное: объект больше не знает никакого шаблона имён и берёт
# всё, что лежит в каталоге.
function Write-Part {
    param(
        [string]   $FileName,
        [int]      $Slot,
        [int]      $Of,
        [string[]] $Body,
        [string]   $PayloadName,
        [switch]   $NoMark,
        [switch]   $NoHeader
    )
    $lines = @()
    if (-not $NoHeader) {
        $lines += ('=== TEST MAILING - PART {0} OF {1} ===' -f $Slot, $Of)
        $lines += ''
        $lines += ('#PART {0:d2} OF {1:d2}' -f $Slot, $Of)
        if ($PayloadName) { $lines += ('#NAME ' + $PayloadName) }
        $lines += ''
    }
    foreach ($b in $Body) { $lines += $(if ($NoMark) { $b } else { '|' + $b }) }
    foreach ($d in (Get-UnpackDirs)) {
        [IO.File]::WriteAllLines((Join-Path $d $FileName), $lines, (New-Object Text.UTF8Encoding($false)))
    }
}

function Split-B64 {
    param([string]$Text, [int]$Chunks)
    $rows = @()
    for ($i = 0; $i -lt $Text.Length; $i += 64) {
        $rows += $Text.Substring($i, [Math]::Min(64, $Text.Length - $i))
    }
    $per = [Math]::Ceiling($rows.Count / $Chunks)
    $out = @()
    for ($c = 0; $c -lt $Chunks; $c++) {
        $take = [Math]::Min($per, $rows.Count - $c * $per)
        if ($take -le 0) { $out += ,@() }
        else { $out += ,@($rows[($c * $per)..($c * $per + $take - 1)]) }
    }
    return $out
}

# Объект не имеет права падать НИ НА ЧЁМ: возвращаемое значение - «прошёл без
# ошибки», и это отдельная проверка в каждом случае, а не техническая обёртка.
function Invoke-Unpacker {
    $threw = $false
    try { Invoke-NAVCodeunit -ServerInstance $ServerInstance -CodeunitId 110207 -CompanyName $CompanyName -ErrorAction Stop }
    catch { $threw = $true }
    return (-not $threw)
}

# --- подготовка --------------------------------------------------------------
if (-not (Test-Path -LiteralPath $admin)) { throw "не найден NavAdminTool.ps1: $admin" }
Import-Module $admin -ErrorAction Stop *>$null

if (-not $CompanyName) {
    $cn = New-Object System.Data.SqlClient.SqlConnection ("Server=localhost;Database=$Database;Integrated Security=True;TrustServerCertificate=True")
    $cn.Open()
    $cmd = $cn.CreateCommand(); $cmd.CommandText = 'SELECT TOP 1 [Name] FROM [Company]'
    $CompanyName = [string]$cmd.ExecuteScalar()
    $cn.Close()
    if (-not $CompanyName) { throw 'в базе нет ни одной компании - задайте -CompanyName' }
}

if ($Import) {
    $tpl = Join-Path $repo 'onsite\Unpacker.template.txt'
    $logDir = Join-Path $env:TEMP 'LineProfiler-unpacker'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    foreach ($step in @(
        @{ Arg = "Command=ImportObjects,File=`"$tpl`",ImportAction=Overwrite,SynchronizeSchemaChanges=Force"; Log = 'import.log' },
        @{ Arg = 'Command=CompileObjects,Filter="Type=Codeunit;ID=110207",SynchronizeSchemaChanges=Force';    Log = 'compile.log' })) {
        $log = Join-Path $logDir $step.Log
        if (Test-Path $log) { Remove-Item $log -Force }
        Start-Process -FilePath $finsql -Wait -NoNewWindow -ArgumentList (
            "$($step.Arg),ServerName=localhost,Database=$Database,Ntauthentication=1,LogFile=`"$log`"") | Out-Null
        if (Test-Path $log) { throw ("finsql: " + (Get-Content $log -Encoding Oem -Raw)) }
    }
    # C/SIDE держит кэш сборки; оставленный процесс мешает следующему прогону.
    Get-Process finsql -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
}

$dirs = @(Get-UnpackDirs)
if ($dirs.Count -eq 0) { throw "не найден рабочий каталог сессий NAV: $navUsers" }
New-Parts

$work = Join-Path $env:TEMP ('lp-unp-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null
$inner = Join-Path $work 'payload.txt'
Set-Content -LiteralPath $inner -Value ((1..40 | ForEach-Object { "line $_ of the payload" }) -join "`r`n") -Encoding Ascii
$zip = Join-Path $work 'payload.zip'
Compress-Archive -Path $inner -DestinationPath $zip
$zipHash = (Get-FileHash $zip -Algorithm SHA256).Hash
$b64     = [Convert]::ToBase64String([IO.File]::ReadAllBytes($zip))
$chunks  = Split-B64 -Text $b64 -Chunks 2

Write-Host ("Стенд: {0}, компания {1}; нагрузка {2} знаков base64" -f $ServerInstance, $CompanyName, $b64.Length)
Write-Host ("Каталог частей: {0}" -f $dirs[0])

# --- 1. две помеченные части по порядку --------------------------------------
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $chunks[0]
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'по порядку: объект не упал' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'по порядку: файл собран' $true ($null -ne $out)
if ($out) {
    Test-Value 'по порядку: тип по сигнатуре' 'zip'     ([IO.Path]::GetExtension($out.Name).TrimStart('.'))
    Test-Value 'по порядку: имя без #NAME'    'payload.zip' $out.Name
    Test-Value 'по порядку: сумма совпала'    $zipHash (Get-FileHash $out.FullName -Algorithm SHA256).Hash
}
Test-Value 'по порядку: склейка по номерам'  $true ((Get-Log) -match 'joined by their own numbers')
Test-Value 'по порядку: про имя сказано'     $true ((Get-Log) -match 'carries no name for the file')

# --- 2. номера обратны именам файлов -----------------------------------------
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 2 -Of 2 -Body $chunks[1]
Write-Part -FileName 'part02.txt' -Slot 1 -Of 2 -Body $chunks[0]
Test-Value 'вперемешку: объект не упал' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'вперемешку: файл собран' $true ($null -ne $out)
if ($out) { Test-Value 'вперемешку: сумма совпала' $zipHash (Get-FileHash $out.FullName -Algorithm SHA256).Hash }

# --- 3. чистый base64 без меток и шапки --------------------------------------
Clear-Parts; Clear-Output
$rows = @()
for ($i = 0; $i -lt $b64.Length; $i += 76) { $rows += $b64.Substring($i, [Math]::Min(76, $b64.Length - $i)) }
Write-Part -FileName 'part01.txt' -Body $rows -NoMark -NoHeader
Test-Value 'чистый base64: объект не упал' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'чистый base64: файл собран' $true ($null -ne $out)
if ($out) { Test-Value 'чистый base64: сумма совпала' $zipHash (Get-FileHash $out.FullName -Algorithm SHA256).Hash }

# --- 4. обрезанная часть, и сразу за ней исправный запуск ---------------------
Clear-Parts; Clear-Output
$cut = @($chunks[0])
$cut[$cut.Count - 1] = $cut[$cut.Count - 1].Substring(0, 61)
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $cut
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'обрезок: объект не упал'     $true (Invoke-Unpacker)
Test-Value 'обрезок: файла не появилось' $true ($null -eq (Get-Output))
Test-Value 'обрезок: причина названа'    $true ((Get-Log) -match 'not a multiple of four')

Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $chunks[0]
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'после отказа: объект не заклинило' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'после отказа: файл собран' $true ($null -ne $out)
if ($out) { Test-Value 'после отказа: сумма совпала' $zipHash (Get-FileHash $out.FullName -Algorithm SHA256).Hash }

# --- 5. посторонний текст ----------------------------------------------------
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -NoMark -NoHeader -Body @(
    'Dear colleagues, please find attached', 'the report for August.')
Test-Value 'чужой файл: объект не упал'     $true (Invoke-Unpacker)
Test-Value 'чужой файл: файла не появилось' $true ($null -eq (Get-Output))
Test-Value 'чужой файл: причина названа'    $true ((Get-Log) -match 'not a plain base64 file')

# --- 6. каталог есть, но пуст ------------------------------------------------
Clear-Parts; Clear-Output
Test-Value 'пустой каталог: объект не упал'     $true (Invoke-Unpacker)
Test-Value 'пустой каталог: файла не появилось' $true ($null -eq (Get-Output))
Test-Value 'пустой каталог: причина названа'    $true ((Get-Log) -match 'There are no files in')

# --- 7. каталога нет вовсе ---------------------------------------------------
# Объект обязан НАЗВАТЬ каталог, который ему нужен, а не молчать: снаружи его никак
# иначе не узнать, а сам он его не создаёт - неудачный запуск не оставляет следов.
Clear-Output
Remove-Parts
Test-Value 'нет каталога: объект не упал'     $true (Invoke-Unpacker)
Test-Value 'нет каталога: файла не появилось' $true ($null -eq (Get-Output))
Test-Value 'нет каталога: каталог назван'     $true ((Get-Log) -match 'unpack')
Test-Value 'нет каталога: сам не создался'    $false (Test-Path -LiteralPath $dirs[0])
New-Parts

# --- 8. имена файлов произвольные --------------------------------------------
# Шаблона имён у объекта больше нет: что лежит в каталоге, то и части.
Clear-Parts; Clear-Output
Write-Part -FileName 'mail from colleague.txt'  -Slot 2 -Of 2 -Body $chunks[1]
Write-Part -FileName 'Fwd re package (1).txt'   -Slot 1 -Of 2 -Body $chunks[0]
Test-Value 'любые имена: объект не упал' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'любые имена: файл собран' $true ($null -ne $out)
if ($out) { Test-Value 'любые имена: сумма совпала' $zipHash (Get-FileHash $out.FullName -Algorithm SHA256).Hash }

# --- 9. имя файла приехало строкой #NAME -------------------------------------
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $chunks[0] -PayloadName 'report.zip'
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'имя из письма: объект не упал' $true (Invoke-Unpacker)
$named = Get-Output -Name 'report.zip'
Test-Value 'имя из письма: файл назван как в письме' $true ($null -ne $named)
Test-Value 'имя из письма: payload.* не появился'    $true ($null -eq (Get-Output))
if ($named) { Test-Value 'имя из письма: сумма совпала' $zipHash (Get-FileHash $named.FullName -Algorithm SHA256).Hash }

# --- 10. в #NAME путь наружу -------------------------------------------------
# Имя приезжает текстом снаружи и становится путём на сервере. Такое имя не чинится
# и не обрезается до последней части - оно отвергается целиком.
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $chunks[0] -PayloadName '..\..\evil.zip'
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'путь в имени: объект не упал' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'путь в имени: имя отвергнуто' 'payload.zip' $(if ($out) { $out.Name } else { '<нет файла>' })
Test-Value 'путь в имени: причина названа' $true ((Get-Log) -match 'cannot be used as a file name')
$escaped = $false
foreach ($d in (Get-NavTempDirs)) {
    $up = Split-Path (Split-Path $d -Parent) -Parent
    if (Test-Path -LiteralPath (Join-Path $up 'evil.zip')) { $escaped = $true }
    if (Test-Path -LiteralPath (Join-Path $d 'evil.zip'))  { $escaped = $true }
}
Test-Value 'путь в имени: наружу ничего не записано' $false $escaped

# --- 11. в #NAME кириллица ---------------------------------------------------
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $chunks[0] -PayloadName 'отчёт.zip'
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'кириллица в имени: объект не упал' $true (Invoke-Unpacker)
$out = Get-Output
Test-Value 'кириллица в имени: имя отвергнуто'  'payload.zip' $(if ($out) { $out.Name } else { '<нет файла>' })
Test-Value 'кириллица в имени: причина названа' $true ((Get-Log) -match 'cannot be used as a file name')

# --- 12. расширение из имени спорит с сигнатурой -----------------------------
# docx - это zip по сигнатуре. Верх берёт имя: отправитель знает, что он послал,
# а сигнатура для этого случая слишком груба.
Clear-Parts; Clear-Output
Write-Part -FileName 'part01.txt' -Slot 1 -Of 2 -Body $chunks[0] -PayloadName 'notes.docx'
Write-Part -FileName 'part02.txt' -Slot 2 -Of 2 -Body $chunks[1]
Test-Value 'имя против сигнатуры: объект не упал' $true (Invoke-Unpacker)
$named = Get-Output -Name 'notes.docx'
Test-Value 'имя против сигнатуры: имя из письма' $true ($null -ne $named)
if ($named) { Test-Value 'имя против сигнатуры: сумма совпала' $zipHash (Get-FileHash $named.FullName -Algorithm SHA256).Hash }

# --- уборка ------------------------------------------------------------------
Clear-Parts; Clear-Output
Remove-Parts
Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ('пройдено {0} из {1}' -f $script:Passed, $script:Total)
foreach ($f in $script:Fails) { Write-Host ('  ' + $f) -ForegroundColor Red }
if ($script:Fails.Count -gt 0) { exit 1 }
Write-Host 'распаковщик base64: без расхождений' -ForegroundColor Green
exit 0
