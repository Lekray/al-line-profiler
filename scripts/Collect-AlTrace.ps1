<#
.SYNOPSIS
    Собирает построчную трассировку C/AL в events.tsv: эфемерная ETW-сессия (Full)
    либо готовые события 705 о долгих SQL из журнала Application (Lite).

.DESCRIPTION
    Full. При EnableFullALFunctionTracing=true платформа шлёт в провайдер
    Microsoft-DynamicsNAV-Server событие на КАЖДЫЙ выполненный оператор C/AL
    (payload: objectType, objectId, functionName, lineNumber, statement), плюс пару
    событий на вход в функцию и выход из неё и события SQL. Скрипт поднимает
    эфемерную сессию ETW, ждёт, пока человек отработает сценарий, гасит сессию и
    раскладывает события в TSV.

        logman create trace <имя> -p "Microsoft-DynamicsNAV-Server" 0xA 4
               -o <файл>.etl -mode Circular -max <МБ> -ets
        ... сценарий ...
        logman stop <имя> -ets

    Почему сессия, а не канал журнала Microsoft-DynamicsNAV-Server/Debug: канал по
    умолчанию 1 МБ и заведён с retention=true, то есть, заполнившись, МОЛЧА перестаёт
    писать; поднять его размер - значит изменить постоянную настройку системы, которая
    переживёт прогон. Сессия с ключом -ets живёт только в памяти, в реестр не
    записывается и исчезает вместе с процессом; канал Debug не трогается вообще.

    Lite. Ничего не включает и ничего не перезапускает: читает из журнала Application
    события 705 (Long running SQL statement), которые платформа пишет сама, если запрос
    вышел за SqlLongRunningThreshold. Прав администратора не требует. Точности до
    оператора не даёт - только сами запросы, их время и таблицы, - но показывает
    узкие места мгновенно и на любой машине.

    Номера событий и порядок полей payload МЕНЯЮТСЯ между версиями платформы, поэтому
    и события, и поля разрешаются по ИМЕНАМ: роль события определяется набором имён
    полей и опкодом, индексы полей берутся из манифеста провайдера и на первой записи
    каждого номера СВЕРЯЮТСЯ с именами из ToXml() (тег Data с атрибутом Name). Если
    манифест и файл разошлись, карта перестраивается по XML. Разобранный манифест
    пишется в manifest.tsv, результат сверки - в collect.log.

    Что скрипт меняет на машине (Full) и как возвращает:
      1. эфемерная сессия ETW      создаётся -> logman stop <имя> -ets;
      2. EnableFullALFunctionTracing в CustomSettings.config -> исходное значение;
      3. служба инстанции          старт/перезапуск -> исходное состояние.
    План печатается ДО первого изменяющего действия. Откат выполняется в блоке finally,
    то есть и при ошибке, и при Ctrl+C; на случай, если процесс убьют, в OutDir лежит
    restore.cmd - тот же откат отдельными командами.

    В базу данных скрипт не обращается вовсе. Зависимостей нет: logman входит в состав
    Windows, System.Diagnostics.Eventing.Reader - в .NET Framework.

    Коды возврата: 0 успех, 1 ошибка, 2 нет прав администратора, 3 нет провайдера ETW,
    4 в манифесте нет события уровня оператора, 5 неизвестное ключевое слово,
    6 не найден CustomSettings.config, 7 не поднялась сессия ETW,
    8 трасса пустая либо потеряна, 9 отказ пользователя в плане изменений.

.PARAMETER Mode
    Full - полная трассировка операторов C/AL через эфемерную сессию ETW; нужны права
    администратора, инстанция перезапускается. Lite - только события 705 о долгих SQL
    из журнала Application: без прав, без включения трассировки, без рестартов.

.PARAMETER ServerInstance
    Имя инстанции NAV. По умолчанию DynamicsNAV110.

.PARAMETER OutDir
    Куда складывать результат. По умолчанию <корень репозитория>\out\trace-<метка времени>.

.PARAMETER SessionName
    Имя эфемерной сессии ETW. По умолчанию LineProfiler-AL.

.PARAMETER MaxSizeMB
    Размер кольцевого файла .etl, МБ. По умолчанию 1024. Кольцо не переполняется, но,
    исчерпав размер, затирает НАЧАЛО трассы - скрипт это обнаруживает и сообщает.

.PARAMETER BufferSizeKB
    Размер буфера сессии, КБ. По умолчанию 1024: событие на каждый оператор C/AL - это
    десятки тысяч записей в секунду, на мелких буферах ядро начинает их терять.

.PARAMETER MinBuffers
    Минимальное число буферов сессии. По умолчанию 16.

.PARAMETER MaxBuffers
    Максимальное число буферов сессии. По умолчанию 64.

.PARAMETER Keywords
    Классы событий провайдера по именам. По умолчанию ALFunctionCallTracing (операторы
    C/AL) и SqlTracing (запросы SQL); маска складывается из манифеста и в NAV 2018
    получается 0xA.

.PARAMETER Calibrate
    Снять цену одного события трассировки: сценарий прогоняется ДВАЖДЫ - сначала без
    трассировки, потом с ней, - и разность календарного времени делится на число
    собранных событий. Никаких объектов в базе для этого не создаётся.

.PARAMETER Yes
    Не спрашивать подтверждения плана изменений. План всё равно печатается.

.PARAMETER KeepTracingOn
    Не откатывать трассировку и службу: пригодится, когда за одним прогоном сразу идёт
    второй. Сессия ETW гасится в любом случае - иначе .etl не дочитать.

.PARAMETER NoRestart
    Не перезапускать инстанцию. EnableFullALFunctionTracing читается службой при старте,
    поэтому без перезапуска событий уровня оператора не будет - скрипт об этом скажет.

.PARAMETER NavServicePath
    Каталог Service установки NAV. По умолчанию берётся из ImagePath службы, а если
    службы нет - ищется в C:\Program Files\Microsoft Dynamics NAV\*\Service.

.PARAMETER ParseFile
    Разобрать ранее собранный .etl (или .evtx) в events.tsv и выйти. Ничего не меняет,
    прав администратора не требует.

.PARAMETER Wait
    Lite: отметить текущий момент, дождаться Enter и взять только те события 705,
    которые платформа записала за это время. Без -Wait берётся история за -Minutes.

.PARAMETER Minutes
    Lite: окно выборки в минутах назад от текущего момента. 0 (по умолчанию) - всё,
    что есть в журнале.

.PARAMETER Since
    Lite: начало окна выборки явной датой-временем. Имеет приоритет над -Minutes.

.PARAMETER Top
    Lite: сколько строк сводки печатать. По умолчанию 12.

.PARAMETER SqlMapFile
    Lite: sqlmap.tsv от Build-KeysIndex.ps1 - словарь "имя таблицы в SQL -> таблица NAV".
    По умолчанию <корень репозитория>\out\sqlmap.tsv, если он есть.

.EXAMPLE
    .\Collect-AlTrace.ps1 -Mode Lite
    Долгие SQL из журнала Application. Прав администратора не требует.

.EXAMPLE
    .\Collect-AlTrace.ps1 -Mode Full -MaxSizeMB 2048
    Полная трассировка операторов, кольцо 2 ГБ. Из консоли администратора.

.EXAMPLE
    .\Collect-AlTrace.ps1 -Mode Full -Calibrate
    То же плюс замер цены события: сценарий выполняется дважды.

.EXAMPLE
    .\Collect-AlTrace.ps1 -ParseFile .\out\trace-20260814-093000\trace.etl
    Перечитать уже собранный файл, ничего не трогая.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [ValidateSet('Full', 'Lite')]
    [string]   $Mode           = 'Full',
    [string]   $ServerInstance = 'DynamicsNAV110',
    [string]   $OutDir,
    [string]   $SessionName    = 'LineProfiler-AL',
    [ValidateRange(16, 16384)]
    [int]      $MaxSizeMB      = 1024,
    [ValidateRange(8, 8192)]
    [int]      $BufferSizeKB   = 1024,
    [ValidateRange(2, 1024)]
    [int]      $MinBuffers     = 16,
    [ValidateRange(2, 1024)]
    [int]      $MaxBuffers     = 64,
    [string[]] $Keywords       = @('ALFunctionCallTracing', 'SqlTracing'),
    [switch]   $Calibrate,
    [switch]   $Yes,
    [switch]   $KeepTracingOn,
    [switch]   $NoRestart,
    [string]   $NavServicePath,
    [string]   $ParseFile,
    [switch]   $Wait,
    [int]      $Minutes        = 0,
    [datetime] $Since,
    [int]      $Top            = 12,
    [string]   $SqlMapFile
)

$ErrorActionPreference = 'Stop'

$PROVIDER    = 'Microsoft-DynamicsNAV-Server'
$SETTING_KEY = 'EnableFullALFunctionTracing'
$LITE_EVENT  = 705                     # Long running SQL statement, журнал Application

# именованные столбцы TSV - эти поля payload раскладываются по колонкам
$NAMED = @('sessionId', 'objectType', 'objectId', 'functionName', 'lineNumber', 'statement')

# поля, которые в сырой столбец не выносим: на каждом событии одинаковы и только
# раздувают файл в разы
$BORING = @{
    'navTenantId'     = $true; 'tenantId'    = $true; 'serverInstanceName' = $true
    'platformVersion' = $true; 'aadTenantId' = $true; 'aadUserId'          = $true
}

# заголовок events.tsv - формат общий для обоих режимов
$TSV_HEADER = @('EventId', 'EventName', 'TimeCreatedTicks', 'SessionId', 'ObjectType', 'ObjectId',
                'FunctionName', 'LineNumber', 'Statement', 'Level', 'RecordId', 'ThreadId', 'Raw')

# ---------------------------------------------------------------------------
# вспомогательное
# ---------------------------------------------------------------------------

$script:Log = $null

function Write-Log {
    param([string] $Message)
    if ($script:Log) { $script:Log.WriteLine(('{0:yyyy-MM-dd HH:mm:ss.fff}  {1}' -f (Get-Date), $Message)) }
    Write-Verbose $Message
}

function Write-Step {
    param([string] $Message)
    Write-Host ('  ' + $Message) -ForegroundColor DarkGray
    Write-Log $Message
}

function Write-Head {
    param([string] $Message)
    Write-Host ''
    Write-Host $Message -ForegroundColor Cyan
    Write-Log ('--- ' + $Message)
}

function Test-Elevated {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Консольные утилиты пишут диагностику в stderr и в кодировке OEM. При
# ErrorActionPreference=Stop слияние 2>&1 превратило бы диагностику в исключение,
# поэтому на время вызова режим смягчается, а кодировка вывода ставится в OEM -
# иначе русские подписи logman не разобрать, а по ним ищется счётчик потерь.
function Invoke-Native {
    param(
        [Parameter(Mandatory)][string] $Exe,
        [string[]] $Arguments,
        [switch]   $AllowFail
    )
    $prevEap = $ErrorActionPreference
    $prevEnc = $null
    try { $prevEnc = [Console]::OutputEncoding } catch { }
    $ErrorActionPreference = 'Continue'
    try {
        try {
            $oemcp = [int](Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\CodePage' -Name OEMCP).OEMCP
            [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding($oemcp)
        } catch { }
        $out  = & $Exe @Arguments 2>&1 | ForEach-Object { [string]$_ }
        $code = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $prevEap
        if ($prevEnc) { try { [Console]::OutputEncoding = $prevEnc } catch { } }
    }

    Write-Log ('{0} {1} -> код {2}' -f $Exe, ($Arguments -join ' '), $code)
    if ($out) { foreach ($l in $out) { if ($l.Trim()) { Write-Log ('    | ' + $l) } } }
    if ($code -ne 0 -and -not $AllowFail) {
        throw ('{0} {1} завершился с кодом {2}: {3}' -f $Exe, ($Arguments -join ' '), $code,
               (@($out | Where-Object { $_.Trim() }) -join ' '))
    }
    return [pscustomobject]@{ Code = $code; Output = @($out) }
}

# Каталог Service установки NAV: сначала из ImagePath службы, затем перебором.
function Resolve-NavServicePath {
    param([string] $Instance, [string] $Hint)
    if ($Hint) {
        if (-not (Test-Path -LiteralPath $Hint)) { throw ('Каталог не найден: ' + $Hint) }
        return (Resolve-Path -LiteralPath $Hint).Path
    }
    $key = 'HKLM:\SYSTEM\CurrentControlSet\Services\MicrosoftDynamicsNavServer$' + $Instance
    $img = $null
    try { $img = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).ImagePath } catch { }
    if ($img -and $img -match '^"([^"]+)"') {
        $dir = Split-Path -Parent $Matches[1]
        if (Test-Path -LiteralPath $dir) { return $dir }
    }
    $probe = @(Get-ChildItem 'C:\Program Files\Microsoft Dynamics NAV' -Directory -ErrorAction SilentlyContinue |
               ForEach-Object { Join-Path $_.FullName 'Service' } |
               Where-Object { Test-Path -LiteralPath $_ } |
               Sort-Object -Descending)
    if ($probe.Count -gt 0) { return $probe[0] }
    return $null
}

# CustomSettings.config: у именованной инстанции он в Instances\<имя>,
# у единственной - прямо в корне Service.
function Resolve-NavConfigPath {
    param([string] $ServiceRoot, [string] $Instance)
    $c = Join-Path $ServiceRoot ('Instances\{0}\CustomSettings.config' -f $Instance)
    if (Test-Path -LiteralPath $c) { return $c }
    $c = Join-Path $ServiceRoot 'CustomSettings.config'
    if (Test-Path -LiteralPath $c) { return $c }
    return $null
}

# ---------------------------------------------------------------------------
# манифест провайдера: имена -> номера событий и индексы полей
# ---------------------------------------------------------------------------

function Get-TraceManifest {
    param([string] $ProviderName)

    $prov = Get-WinEvent -ListProvider $ProviderName -ErrorAction Stop

    # маска ключевых слов беззнаковая: у неё занят старший бит, и в int64 такое
    # значение стало бы отрицательным
    $kw = @{}
    foreach ($k in $prov.Keywords) {
        if ($k.Name -and [int64]$k.Value -ge 0) { $kw[$k.Name] = [uint64]$k.Value }
    }

    $rxField = New-Object System.Text.RegularExpressions.Regex('<data\s+name\s*=\s*"([^"]+)"', 'IgnoreCase')

    $byId = @{}
    foreach ($e in $prov.Events) {
        $id = [int]$e.Id
        # один номер может присутствовать в нескольких версиях - берём старшую
        if ($byId.ContainsKey($id) -and $byId[$id].Version -ge [int]$e.Version) { continue }

        $names = New-Object System.Collections.Generic.List[string]
        if ($e.Template) {
            foreach ($m in $rxField.Matches($e.Template)) { [void]$names.Add($m.Groups[1].Value) }
        }
        $idx = @{}
        for ($i = 0; $i -lt $names.Count; $i++) { $idx[$names[$i]] = $i }

        $opName = ''; $opValue = 0
        if ($e.Opcode) { $opName = [string]$e.Opcode.Name; $opValue = [int]$e.Opcode.Value }
        $taskName = ''
        if ($e.Task) { $taskName = [string]$e.Task.Name }
        $kwNames = @(); $kwMaskEvent = [uint64]0
        foreach ($k in $e.Keywords) {
            if ($k.Name) { $kwNames += $k.Name }
            if ([int64]$k.Value -ge 0) { $kwMaskEvent = $kwMaskEvent -bor [uint64]$k.Value }
        }
        $channel = ''
        if ($e.LogLink) { $channel = [string]$e.LogLink.LogName }
        $levelName = ''; $levelValue = 0
        if ($e.Level) { $levelName = [string]$e.Level.Name; $levelValue = [int]$e.Level.Value }

        # роль определяется набором имён полей и опкодом; номер события не участвует
        $role = ''
        if ($idx.ContainsKey('lineNumber') -and $idx.ContainsKey('statement')) {
            $role = 'ALStatement'
        }
        elseif ($idx.ContainsKey('failureMessage') -and $idx.ContainsKey('functionName')) {
            $role = 'ALFunctionError'
        }
        elseif ($idx.ContainsKey('functionName') -and $idx.ContainsKey('objectId')) {
            if     ($opValue -eq 1) { $role = 'ALFunctionStart' }
            elseif ($opValue -eq 2) { $role = 'ALFunctionStop' }
            else                    { $role = 'ALFunction' }
        }
        elseif ($idx.ContainsKey('sqlStatement')) {
            $suffix = ''
            if     ($opValue -eq 1) { $suffix = ':Start' }
            elseif ($opValue -eq 2) { $suffix = ':Stop' }
            $role = 'Sql:' + $taskName + $suffix
        }
        else {
            $role = $taskName
            if ($opName) { $role = $role + ':' + ($opName -replace '^win:', '') }
            if (-not $role.Trim(':')) { $role = 'Event' + $id }
        }

        $byId[$id] = @{
            Id       = $id;     Version     = [int]$e.Version; Role     = $role;  Task = $taskName
            Opcode   = $opName; OpcodeValue = $opValue;        Level    = $levelName
            LevelValue = $levelValue
            Keywords = ($kwNames -join '+');  KeywordMask = $kwMaskEvent
            Channel  = $channel
            Fields   = $idx;    FieldNames  = $names.ToArray()
        }
    }
    return @{ Provider = $prov; Events = $byId; Keywords = $kw }
}

function Get-FieldIndex {
    param($Info, [string] $Name)
    if ($Info.Fields.ContainsKey($Name)) { return [int]$Info.Fields[$Name] }
    return -1
}

function Write-ManifestTsv {
    param([hashtable] $Manifest, [string] $Path)
    $rows = New-Object System.Collections.Generic.List[string]
    [void]$rows.Add((@('EventId', 'Version', 'Role', 'Task', 'Opcode', 'OpcodeValue', 'Level', 'LevelValue',
                       'Keywords', 'KeywordMask', 'Channel', 'FieldCount', 'Fields', 'IdxSessionId',
                       'IdxObjectType', 'IdxObjectId', 'IdxFunctionName', 'IdxLineNumber',
                       'IdxStatement') -join "`t"))
    foreach ($id in ($Manifest.Events.Keys | Sort-Object)) {
        $i = $Manifest.Events[$id]
        [void]$rows.Add((@(
            $i.Id, $i.Version, $i.Role, $i.Task, $i.Opcode, $i.OpcodeValue, $i.Level, $i.LevelValue
            $i.Keywords, ('0x{0:X}' -f $i.KeywordMask), $i.Channel
            $i.FieldNames.Count, ($i.FieldNames -join ',')
            (Get-FieldIndex $i 'sessionId'), (Get-FieldIndex $i 'objectType'), (Get-FieldIndex $i 'objectId')
            (Get-FieldIndex $i 'functionName'), (Get-FieldIndex $i 'lineNumber'), (Get-FieldIndex $i 'statement')
        ) -join "`t"))
    }
    [System.IO.File]::WriteAllText($Path, (($rows -join "`r`n") + "`r`n"),
        (New-Object System.Text.UTF8Encoding($false)))
}

# ---------------------------------------------------------------------------
# разбор собранного файла (.etl или .evtx) в events.tsv
# ---------------------------------------------------------------------------
# Вынесено в функцию, чтобы разбор можно было прогнать отдельно от сбора: сбор
# требует прав администратора, разбор нет (ключ -ParseFile).

function ConvertTo-TraceTsv {
    param(
        [hashtable] $Manifest,
        [string]    $TracePath,
        [string]    $TsvPath,
        [string]    $ProviderName
    )

    # Индексы полей раскладываются заранее: в цикле на миллион записей поиск по
    # хеш-таблице и вызовы функций стоят дороже самого разбора. Карта строится по
    # ИМЕНАМ полей из манифеста и на первой записи каждого номера сверяется с
    # именами из XML самой записи - см. ниже $verified.
    $fast = @{}
    foreach ($id in $Manifest.Events.Keys) {
        $fast[$id] = New-FieldMap -FieldNames $Manifest.Events[$id].FieldNames -Role $Manifest.Events[$id].Role
    }

    # в TSV не должно быть ни табуляций, ни переносов: оператор C/AL бывает многострочным
    $rxWs   = New-Object System.Text.RegularExpressions.Regex('[\t\r\n]+', 'Compiled')
    $rxData = New-Object System.Text.RegularExpressions.Regex(
        '<Data\s+Name\s*=\s*[''"]([^''"]+)[''"]', ([System.Text.RegularExpressions.RegexOptions]'IgnoreCase,Compiled'))

    $writer = New-Object System.IO.StreamWriter($TsvPath, $false, (New-Object System.Text.UTF8Encoding($false)))
    $writer.NewLine = "`r`n"
    $writer.WriteLine(($TSV_HEADER -join "`t"))

    $byRole     = @{}
    $verified   = @{}          # номер события -> уже сверяли имена полей
    $verifyNote = New-Object System.Collections.Generic.List[string]
    $foreign    = 0
    $unresolved = 0
    $total      = 0
    $tickMin    = [int64]::MaxValue
    $tickMax    = [int64]::MinValue

    $query = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery(
        $TracePath, [System.Diagnostics.Eventing.Reader.PathType]::FilePath)
    $query.ReverseDirection = $false          # от старых к новым, как -Oldest у Get-WinEvent
    $reader = New-Object System.Diagnostics.Eventing.Reader.EventLogReader($query)
    try {
        while ($true) {
            $rec = $reader.ReadEvent()
            if ($null -eq $rec) { break }
            try {
                $total++
                $id = [int]$rec.Id
                $isOurs = ($rec.ProviderName -eq $ProviderName)
                if (-not $isOurs) {
                    # пустое имя провайдера в .etl значит, что манифест не подставился
                    # и payload остался сырым: это диагноз, а не мелочь
                    if ([string]::IsNullOrEmpty($rec.ProviderName)) { $unresolved++ } else { $foreign++ }
                }

                $ticks = 0L
                if ($null -ne $rec.TimeCreated) {
                    $ticks = $rec.TimeCreated.Ticks
                    if ($ticks -lt $tickMin) { $tickMin = $ticks }
                    if ($ticks -gt $tickMax) { $tickMax = $ticks }
                }

                # карта полей применима ТОЛЬКО к событиям своего провайдера: номера
                # событий у разных провайдеров совпадают сплошь и рядом, и чужая
                # запись с тем же номером молча растащилась бы по чужим колонкам
                $f = $null
                if ($isOurs) { $f = $fast[$id] }

                # сверка имён: один раз на каждый номер события. Имена в XML идут в том
                # же порядке, что значения в Properties, - это и есть чтение по именам.
                if ($null -ne $f -and -not $verified.ContainsKey($id)) {
                    $verified[$id] = $true
                    $xmlNames = @()
                    try {
                        foreach ($m in $rxData.Matches($rec.ToXml())) { $xmlNames += $m.Groups[1].Value }
                    } catch { }
                    if ($xmlNames.Count -gt 0) {
                        if (($xmlNames -join ',') -ne ($f.Names -join ',')) {
                            $f = New-FieldMap -FieldNames $xmlNames -Role $f.Role
                            $fast[$id] = $f
                            [void]$verifyNote.Add(('событие {0}: манифест разошёлся с файлом, карта перестроена по XML: {1}' -f $id, ($xmlNames -join ',')))
                        } else {
                            [void]$verifyNote.Add(('событие {0}: имена полей сверены с XML, совпали ({1})' -f $id, ($xmlNames -join ',')))
                        }
                    } else {
                        [void]$verifyNote.Add(('событие {0}: в XML нет именованных полей, карта взята из манифеста' -f $id))
                    }
                }

                $props = $rec.Properties
                $pc    = $props.Count
                $raw   = New-Object System.Text.StringBuilder

                if ($null -eq $f) {
                    # чужой провайдер либо события нет в манифесте - payload целиком в сырой столбец
                    $name = 'Event' + $id
                    for ($k = 0; $k -lt $pc; $k++) {
                        if ($k -gt 0) { [void]$raw.Append('; ') }
                        [void]$raw.Append($k).Append('=').Append([string]$props[$k].Value)
                    }
                    $row = @($id, $name, $ticks, '', '', '', '', '', '',
                             [int]$rec.Level, [string]$rec.RecordId, [string]$rec.ThreadId,
                             $rxWs.Replace($raw.ToString(), ' '))
                } else {
                    $name = $f.Role
                    $ri   = $f.RawIdx
                    for ($k = 0; $k -lt $ri.Length; $k++) {
                        $j = $ri[$k]
                        if ($j -ge $pc) { continue }
                        if ($raw.Length -gt 0) { [void]$raw.Append('; ') }
                        [void]$raw.Append($f.RawName[$k]).Append('=').Append([string]$props[$j].Value)
                    }
                    $vSes = ''; $vOt = ''; $vOid = ''; $vFn = ''; $vLn = ''; $vSt = ''
                    $k = $f.Ses; if ($k -ge 0 -and $k -lt $pc) { $vSes = [string]$props[$k].Value }
                    $k = $f.OT;  if ($k -ge 0 -and $k -lt $pc) { $vOt  = [string]$props[$k].Value }
                    $k = $f.OID; if ($k -ge 0 -and $k -lt $pc) { $vOid = [string]$props[$k].Value }
                    $k = $f.FN;  if ($k -ge 0 -and $k -lt $pc) { $vFn  = [string]$props[$k].Value }
                    $k = $f.LN;  if ($k -ge 0 -and $k -lt $pc) { $vLn  = [string]$props[$k].Value }
                    $k = $f.ST;  if ($k -ge 0 -and $k -lt $pc) { $vSt  = [string]$props[$k].Value }
                    $row = @($id, $name, $ticks, $vSes, $vOt, $vOid, $vFn, $vLn,
                             $rxWs.Replace($vSt, ' '),
                             [int]$rec.Level, [string]$rec.RecordId, [string]$rec.ThreadId,
                             $rxWs.Replace($raw.ToString(), ' '))
                }

                $writer.WriteLine([string]::Join("`t", $row))
                if ($byRole.ContainsKey($name)) { $byRole[$name]++ } else { $byRole[$name] = 1 }

                if (($total % 200000) -eq 0) {
                    Write-Host ('  разобрано {0:N0}...' -f $total) -ForegroundColor DarkGray
                }
            }
            finally { $rec.Dispose() }
        }
    }
    finally {
        $reader.Dispose()
        $writer.Flush(); $writer.Dispose()
    }

    return @{
        Total = $total; ByRole = $byRole; Foreign = $foreign; Unresolved = $unresolved
        TickMin = $tickMin; TickMax = $tickMax; Verify = $verifyNote
    }
}

# Карта "имя поля -> индекс в Properties" для одного номера события.
function New-FieldMap {
    param([string[]] $FieldNames, [string] $Role)
    $idx = @{}
    for ($i = 0; $i -lt $FieldNames.Count; $i++) { $idx[$FieldNames[$i]] = $i }
    $get = {
        param($n)
        if ($idx.ContainsKey($n)) { return [int]$idx[$n] }
        return -1
    }
    $rawIdx  = New-Object System.Collections.Generic.List[int]
    $rawName = New-Object System.Collections.Generic.List[string]
    for ($k = 0; $k -lt $FieldNames.Count; $k++) {
        $n = $FieldNames[$k]
        if ($NAMED -contains $n) { continue }
        if ($BORING.ContainsKey($n)) { continue }
        [void]$rawIdx.Add($k); [void]$rawName.Add($n)
    }
    return @{
        Role = $Role
        Names = @($FieldNames)
        Ses = (& $get 'sessionId');    OT = (& $get 'objectType')
        OID = (& $get 'objectId');     FN = (& $get 'functionName')
        LN  = (& $get 'lineNumber');   ST = (& $get 'statement')
        RawIdx = $rawIdx.ToArray();    RawName = $rawName.ToArray()
    }
}

# ---------------------------------------------------------------------------
# режим Lite: события 705 о долгих SQL из журнала Application
# ---------------------------------------------------------------------------

# Полезная нагрузка 705 - одна строка вида "ключ: значение" плюс сам текст запроса
# между строками "Connection ID:" и "ProcessId:".
function ConvertFrom-LongRunningSql {
    param([string] $Text)

    $lines = $Text -split "`r`n|`n"
    $sql   = New-Object System.Collections.Generic.List[string]
    $inSql = $false
    foreach ($ln in $lines) {
        if ($ln -match '^\s*Connection ID:') { $inSql = $true; continue }
        if ($ln -match '^ProcessId:')        { $inSql = $false; continue }
        if ($inSql) { $t = $ln.Trim(); if ($t) { [void]$sql.Add($t) } }
    }

    $get = {
        param($rx)
        if ($Text -match $rx) { return $Matches[1].Trim() }
        return ''
    }
    return [pscustomobject]@{
        Category   = (& $get '(?m)^Category:\s*(.+)$')
        SessionId  = (& $get '(?m)^ClientSessionId:\s*(\S+)')
        DurationMs = (& $get '(?m)Execution time:\s*(\d+)\s*ms')
        ThresholdMs= (& $get '(?m)Threshold:\s*(\d+)\s*ms')
        Message    = (& $get '(?m)^\s*Message:\s*(.+)$')
        TaskId     = (& $get '(?m)Task ID:\s*(\S+)')
        ConnId     = (& $get '(?m)Connection ID:\s*(\S+)')
        Tag        = (& $get '(?m)^Tag:\s*(\S+)')
        ThreadId   = (& $get '(?m)^ThreadId:\s*(\S+)')
        ProcessId  = (& $get '(?m)^ProcessId:\s*(\S+)')
        Sql        = ($sql -join ' ')
    }
}

# Форма запроса: литералы и параметры схлопнуты, чтобы одинаковые по смыслу
# запросы сложились в одну строку сводки.
function Get-SqlShape {
    param([string] $Sql)
    $s = $Sql
    $s = [regex]::Replace($s, "'(?:[^']|'')*'", "'?'")
    $s = [regex]::Replace($s, '@[A-Za-z_]\w*', '@?')
    $s = [regex]::Replace($s, '(?<![\w\[$])\d+(\.\d+)?', '?')
    $s = [regex]::Replace($s, '\s+', ' ')
    return $s.Trim()
}

# Имена таблиц из запроса. В SQL таблица компании называется "<Компания>$<Таблица>",
# общая - просто "<Таблица>"; для сопоставления с NAV берётся часть после последнего $.
function Get-SqlTables {
    param([string] $Sql, [hashtable] $SqlMap)
    $res = New-Object System.Collections.Generic.List[string]
    # имя таблицы может быть с квалификаторами и в любом сочетании скобок:
    # FROM [NAV].[dbo].[Компания$Таблица], UPDATE [NAV].dbo.[Таблица], JOIN [Таблица]
    foreach ($m in [regex]::Matches($Sql, '(?i)(?:FROM|JOIN|UPDATE|INTO)\s+(?:(?:\[[^\]]+\]|\w+)\.){0,3}\[([^\]]+)\]')) {
        $raw = $m.Groups[1].Value
        $bare = $raw
        $p = $bare.LastIndexOf('$')
        if ($p -ge 0 -and $p -lt ($bare.Length - 1)) { $bare = $bare.Substring($p + 1) }
        $shown = $bare
        if ($SqlMap -and $SqlMap.ContainsKey($bare)) { $shown = $SqlMap[$bare] }
        if (-not $res.Contains($shown)) { [void]$res.Add($shown) }
    }
    return $res.ToArray()
}

function Import-SqlMap {
    param([string] $Path)
    $map = @{}
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $map }
    $lines = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    for ($i = 1; $i -lt $lines.Length; $i++) {
        $c = $lines[$i] -split "`t"
        if ($c.Length -ge 3 -and $c[0]) { $map[$c[0]] = ('{0} ({1})' -f $c[2], $c[1]) }
    }
    return $map
}

# ---------------------------------------------------------------------------
# служба инстанции
# ---------------------------------------------------------------------------

function Set-NavServiceRunning {
    param($Service, [switch] $Restart)
    if (-not $Service) { return }
    $Service.Refresh()
    if ($Restart -and $Service.Status -eq 'Running') {
        Write-Step ('останов службы {0}...' -f $Service.Name)
        Stop-Service -Name $Service.Name -Force
        $Service.WaitForStatus('Stopped', [TimeSpan]::FromMinutes(5))
    }
    $Service.Refresh()
    if ($Service.Status -ne 'Running') {
        Write-Step ('запуск службы {0}...' -f $Service.Name)
        Start-Service -Name $Service.Name
        $Service.WaitForStatus('Running', [TimeSpan]::FromMinutes(5))
    }
    $Service.Refresh()
    Write-Step ('служба {0}: {1}' -f $Service.Name, $Service.Status)
}

function Stop-NavService {
    param($Service)
    if (-not $Service) { return }
    $Service.Refresh()
    if ($Service.Status -ne 'Stopped') {
        Stop-Service -Name $Service.Name -Force
        $Service.WaitForStatus('Stopped', [TimeSpan]::FromMinutes(5))
    }
    $Service.Refresh()
    Write-Step ('служба {0}: {1}' -f $Service.Name, $Service.Status)
}

# ---------------------------------------------------------------------------
# итог по разобранной трассе (общий для -ParseFile и режима Full)
# ---------------------------------------------------------------------------

function Write-TraceSummary {
    param(
        [hashtable] $Stats,
        [string]    $TracePath,
        [double]    $ScenarioSec = 0
    )
    $total = $Stats.Total
    $span  = 0.0
    if ($total -gt 0 -and $Stats.TickMax -ge $Stats.TickMin) {
        $span = ($Stats.TickMax - $Stats.TickMin) / 10000000.0
    }
    $sizeMB = 0.0
    if (Test-Path -LiteralPath $TracePath) { $sizeMB = (Get-Item -LiteralPath $TracePath).Length / 1MB }

    Write-Host ''
    Write-Host ('Файл:      {0} ({1:N1} МБ)' -f (Split-Path -Leaf $TracePath), $sizeMB)
    Write-Host ('Событий:   {0:N0}' -f $total)
    foreach ($k in ($Stats.ByRole.Keys | Sort-Object { -$Stats.ByRole[$_] } | Select-Object -First 12)) {
        Write-Host ('  {0,-28} {1,10:N0}' -f $k, $Stats.ByRole[$k])
    }
    if ($ScenarioSec -gt 0) {
        Write-Host ('Сбор:      {0:N1} с сценарий, {1:N1} с между первым и последним событием' -f $ScenarioSec, $span)
    } else {
        Write-Host ('Охват:     {0:N1} с между первым и последним событием' -f $span)
    }
    if ($Stats.Unresolved -gt 0) {
        Write-Host ('Не разобрано манифестом: {0:N0} записей - payload остался сырым' -f $Stats.Unresolved) -ForegroundColor Yellow
    }
    if ($Stats.Foreign -gt 0) {
        Write-Host ('Записей чужих провайдеров: {0:N0}' -f $Stats.Foreign) -ForegroundColor DarkGray
    }
}

# ---------------------------------------------------------------------------
# старт
# ---------------------------------------------------------------------------

if ($PSVersionTable.PSEdition -eq 'Core') {
    Write-Host 'Нужен Windows PowerShell 5.1 (powershell.exe), а не PowerShell Core.' -ForegroundColor Red
    exit 1
}

if (-not $OutDir) {
    $taskRoot = Split-Path -Parent $PSScriptRoot          # scripts -> LineProfiler
    $OutDir   = Join-Path $taskRoot ('out\trace-{0:yyyyMMdd-HHmmss}' -f (Get-Date))
}
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$utf8NoBom  = New-Object System.Text.UTF8Encoding($false)
$script:Log = New-Object System.IO.StreamWriter((Join-Path $OutDir 'collect.log'), $false, $utf8NoBom)
$script:Log.AutoFlush = $true          # чтобы журнал уцелел, даже если процесс убьют
$script:Log.NewLine   = "`r`n"

$tsvPath  = Join-Path $OutDir 'events.tsv'
$etlPath  = Join-Path $OutDir 'trace.etl'
# Коды возврата: 0 - сбор прошёл; 8 - трассе верить нельзя (событий нет, кольцо
# перезаписало начало либо ядро сообщило о потерянных буферах); 9 - откат не удался,
# машина осталась изменённой и требует рук. Девятка перебивает восьмёрку: сперва
# приводят в порядок сервер, потом разбираются с данными.
$exitCode = 0

Write-Host ''
Write-Host ('Сбор трассировки C/AL: режим {0}, инстанция {1}' -f $Mode, $ServerInstance)
Write-Host ('Каталог:   {0}' -f $OutDir)
Write-Log ('=== старт; режим {0}; инстанция {1}; сессия {2}; кольцо {3} МБ; Calibrate={4} ===' -f
    $Mode, $ServerInstance, $SessionName, $MaxSizeMB, [bool]$Calibrate)

# --- манифест провайдера ----------------------------------------------------
# нужен режиму Full и разбору файла; в Lite он не обязателен, но manifest.tsv
# пишется всё равно - по нему видно, что вообще умеет платформа на этой машине
$manifest = $null
try { $manifest = Get-TraceManifest -ProviderName $PROVIDER }
catch {
    Write-Log ('манифест не прочитан: ' + $_.Exception.Message)
    if ($Mode -eq 'Full' -or $ParseFile) {
        Write-Host ''
        Write-Host ('Провайдер ETW {0} не зарегистрирован на этой машине.' -f $PROVIDER) -ForegroundColor Red
        Write-Host 'Он ставится вместе с сервером NAV (Microsoft.Dynamics.Nav.Ncl.etw.dll).'
        $script:Log.Dispose()
        exit 3
    }
    Write-Step 'манифест провайдера недоступен - для режима Lite это не важно'
}
if ($manifest) {
    Write-ManifestTsv -Manifest $manifest -Path (Join-Path $OutDir 'manifest.tsv')
    Write-Step ('манифест провайдера разобран: событий {0}, manifest.tsv записан' -f $manifest.Events.Count)
}

# ===========================================================================
# разбор готового файла: ничего не меняет, прав не требует
# ===========================================================================
if ($ParseFile) {
    if (-not (Test-Path -LiteralPath $ParseFile)) {
        Write-Host ('Файл не найден: ' + $ParseFile) -ForegroundColor Red
        $script:Log.Dispose()
        exit 1
    }
    $ParseFile = (Resolve-Path -LiteralPath $ParseFile).Path
    Write-Head ('Разбор файла ' + $ParseFile)
    $stats = ConvertTo-TraceTsv -Manifest $manifest -TracePath $ParseFile -TsvPath $tsvPath -ProviderName $PROVIDER
    foreach ($v in $stats.Verify) { Write-Log ('сверка полей: ' + $v) }
    Write-TraceSummary -Stats $stats -TracePath $ParseFile
    Write-Host ('Результат: {0}' -f $tsvPath)
    if ($stats.Total -eq 0) { $exitCode = 8 }
    $script:Log.Flush(); $script:Log.Dispose(); $script:Log = $null
    exit $exitCode
}

# ===========================================================================
# режим Lite: события 705 из журнала Application
# ===========================================================================
if ($Mode -eq 'Lite') {
    $liteProvider = 'MicrosoftDynamicsNavServer$' + $ServerInstance
    $start = $null

    if ($Wait) {
        $start = Get-Date
        Write-Host ''
        Write-Host '================================================================' -ForegroundColor Green
        Write-Host ' РЕЖИМ LITE: настройки машины НЕ меняются' -ForegroundColor Green
        Write-Host '================================================================' -ForegroundColor Green
        Write-Host ' Платформа сама пишет событие 705, если запрос вышел за порог'
        Write-Host ' SqlLongRunningThreshold. Выполните сценарий в NAV.'
        Write-Host ''
        Read-Host ' Нажмите Enter, когда сценарий отработает' | Out-Null
    }
    elseif ($PSBoundParameters.ContainsKey('Since')) { $start = $Since }
    elseif ($Minutes -gt 0)                          { $start = (Get-Date).AddMinutes(-$Minutes) }

    $filter = @{ LogName = 'Application'; ProviderName = $liteProvider; Id = $LITE_EVENT }
    if ($start) { $filter['StartTime'] = $start }
    Write-Head ('Чтение журнала Application, событие {0}, провайдер {1}' -f $LITE_EVENT, $liteProvider)
    if ($start) { Write-Step ('окно с {0:yyyy-MM-dd HH:mm:ss}' -f $start) } else { Write-Step 'окно: весь журнал' }

    $events = @()
    try { $events = @(Get-WinEvent -FilterHashtable $filter -ErrorAction Stop) }
    catch {
        Write-Log ('Get-WinEvent: ' + $_.Exception.Message)
        Write-Step ('событий 705 в окне нет (' + $_.Exception.Message + ')')
    }
    Write-Step ('получено записей: {0}' -f $events.Count)

    if (-not $SqlMapFile) {
        $probe = Join-Path (Split-Path -Parent $PSScriptRoot) 'out\sqlmap.tsv'
        if (Test-Path -LiteralPath $probe) { $SqlMapFile = $probe }
    }
    $sqlMap = Import-SqlMap -Path $SqlMapFile
    if ($sqlMap.Count -gt 0) { Write-Step ('словарь таблиц: {0} имён из {1}' -f $sqlMap.Count, $SqlMapFile) }

    # события пишутся в тот же events.tsv: колонки времени оператора остаются
    # пустыми, длительность и порог уходят в сырой столбец
    $rxWs   = New-Object System.Text.RegularExpressions.Regex('[\t\r\n]+', 'Compiled')
    $writer = New-Object System.IO.StreamWriter($tsvPath, $false, $utf8NoBom)
    $writer.NewLine = "`r`n"
    $writer.WriteLine(($TSV_HEADER -join "`t"))

    $agg     = @{}
    $totalMs = 0L
    $maxMs   = 0
    $parsed  = 0
    try {
        foreach ($ev in ($events | Sort-Object TimeCreated)) {
            $text = ''
            if ($ev.Properties.Count -gt 0) { $text = [string]$ev.Properties[0].Value }
            $p = ConvertFrom-LongRunningSql -Text $text
            if (-not $p.Sql) { continue }
            $parsed++
            $ms = 0
            if ($p.DurationMs) { $ms = [int]$p.DurationMs }
            $totalMs += $ms
            if ($ms -gt $maxMs) { $maxMs = $ms }

            $raw = 'durationMs={0}; thresholdMs={1}; category={2}; taskId={3}; connectionId={4}; tag={5}; processId={6}' -f
                   $p.DurationMs, $p.ThresholdMs, $p.Category, $p.TaskId, $p.ConnId, $p.Tag, $p.ProcessId
            $row = @($ev.Id, 'SlowSql', $ev.TimeCreated.Ticks, $p.SessionId, '', '', '', '',
                     $rxWs.Replace($p.Sql, ' '), [int]$ev.Level, [string]$ev.RecordId, $p.ThreadId, $raw)
            $writer.WriteLine([string]::Join("`t", $row))

            $shape = Get-SqlShape -Sql $p.Sql
            if ($shape.Length -gt 3000) { $shape = $shape.Substring(0, 3000) }
            if (-not $agg.ContainsKey($shape)) {
                $agg[$shape] = [pscustomobject]@{
                    Count = 0; TotalMs = 0L; MaxMs = 0; Tables = (Get-SqlTables -Sql $p.Sql -SqlMap $sqlMap)
                    Shape = $shape
                }
            }
            $a = $agg[$shape]
            $a.Count   = $a.Count + 1
            $a.TotalMs = $a.TotalMs + $ms
            if ($ms -gt $a.MaxMs) { $a.MaxMs = $ms }
        }
    }
    finally { $writer.Flush(); $writer.Dispose() }

    # сводка по формам запросов
    $ranked = @($agg.Values | Sort-Object -Property TotalMs -Descending)
    $rows = New-Object System.Collections.Generic.List[string]
    [void]$rows.Add((@('Rank', 'Count', 'TotalMs', 'AvgMs', 'MaxMs', 'Tables', 'Shape') -join "`t"))
    for ($i = 0; $i -lt $ranked.Count; $i++) {
        $a = $ranked[$i]
        $avg = 0; if ($a.Count -gt 0) { $avg = [int]($a.TotalMs / $a.Count) }
        [void]$rows.Add((@(($i + 1), $a.Count, $a.TotalMs, $avg, $a.MaxMs, ($a.Tables -join ', '), $a.Shape) -join "`t"))
    }
    $slowPath = Join-Path $OutDir 'slowsql.tsv'
    [System.IO.File]::WriteAllText($slowPath, (($rows -join "`r`n") + "`r`n"), $utf8NoBom)

    Write-Host ''
    Write-Host ('Событий 705: {0:N0}, разобрано {1:N0}, суммарно {2:N1} с, максимум {3:N1} с' -f
        $events.Count, $parsed, ($totalMs / 1000.0), ($maxMs / 1000.0))
    if ($parsed -gt 0) {
        Write-Host ''
        Write-Host ('{0,4} {1,6} {2,9} {3,8} {4,-26} {5}' -f '№', 'Раз', 'Всего,мс', 'Макс,мс', 'Таблица', 'Запрос')
        Write-Host ('---- ------ --------- -------- -------------------------- ' + ('-' * 44))
        $show = [Math]::Min([Math]::Max($Top, 1), $ranked.Count)
        for ($i = 0; $i -lt $show; $i++) {
            $a = $ranked[$i]
            $tb = ($a.Tables | Select-Object -First 1)
            if (-not $tb) { $tb = '-' }
            if ($tb.Length -gt 26) { $tb = $tb.Substring(0, 23) + '...' }
            $sh = $a.Shape
            if ($sh.Length -gt 44) { $sh = $sh.Substring(0, 41) + '...' }
            Write-Host ('{0,4} {1,6:N0} {2,9:N0} {3,8:N0} {4,-26} {5}' -f ($i + 1), $a.Count, $a.TotalMs, $a.MaxMs, $tb, $sh)
        }
    }
    Write-Host ''
    Write-Host ('Файлы:     events.tsv, slowsql.tsv, manifest.tsv, collect.log')
    Write-Host ('Каталог:   {0}' -f $OutDir)
    Write-Host 'Машина не изменялась: трассировка не включалась, служба не перезапускалась.' -ForegroundColor Green
    Write-Log ('итог Lite: событий {0}, разобрано {1}, форм {2}, суммарно {3} мс' -f
        $events.Count, $parsed, $ranked.Count, $totalMs)

    if ($parsed -eq 0) {
        Write-Host ''
        Write-Host 'Событий 705 в окне нет. Это НЕ ошибка: их пишет сама платформа и только' -ForegroundColor Yellow
        Write-Host 'когда запрос вышел за SqlLongRunningThreshold. Расширьте окно (-Minutes 0)' -ForegroundColor Yellow
        Write-Host 'либо снимайте полную трассировку: -Mode Full.' -ForegroundColor Yellow
        $exitCode = 8
    }
    $script:Log.Flush(); $script:Log.Dispose(); $script:Log = $null
    exit $exitCode
}

# ===========================================================================
# режим Full: эфемерная сессия ETW
# ===========================================================================

# --- права ------------------------------------------------------------------
# отказ внятный: создать сессию ETW и переписать CustomSettings.config без
# повышения нельзя, и падать на середине с чужим исключением незачем
if (-not (Test-Elevated)) {
    Write-Host ''
    Write-Host 'Нужны права администратора: режим Full меняет состояние машины.' -ForegroundColor Red
    Write-Host '  - создаёт эфемерную сессию ETW (logman ... -ets)'
    Write-Host ('  - переключает {0} в CustomSettings.config' -f $SETTING_KEY)
    Write-Host ('  - перезапускает службу MicrosoftDynamicsNavServer${0}' -f $ServerInstance)
    Write-Host ''
    Write-Host 'Запустите в консоли, поднятой от имени администратора:'
    Write-Host ('  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $PSCommandPath) -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Либо снимите долгие SQL без прав вообще:'
    Write-Host ('  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Mode Lite' -f $PSCommandPath) -ForegroundColor Yellow
    Write-Log 'отказ: нет прав администратора'
    $script:Log.Dispose()
    exit 2
}

# --- событие уровня оператора -----------------------------------------------
$stmtEvents = @($manifest.Events.Values | Where-Object { $_.Role -eq 'ALStatement' } | Sort-Object { $_.Id })
if ($stmtEvents.Count -eq 0) {
    Write-Host ''
    Write-Host 'В манифесте провайдера нет события уровня оператора C/AL.' -ForegroundColor Red
    Write-Host 'Искали событие, в payload которого есть поля lineNumber и statement. Что есть:'
    foreach ($id in ($manifest.Events.Keys | Sort-Object)) {
        $i = $manifest.Events[$id]
        Write-Host ('  {0,-6} {1,-26} {2}' -f $i.Id, $i.Role, ($i.FieldNames -join ','))
    }
    Write-Host ('Полная карта: {0}' -f (Join-Path $OutDir 'manifest.tsv'))
    Write-Log 'события ALStatement в манифесте нет - отказ'
    $script:Log.Dispose()
    exit 4
}
$stmt = $stmtEvents[0]
Write-Step ('событие оператора: {0} Id={1} v{2}' -f $stmt.Role, $stmt.Id, $stmt.Version)
Write-Step ('поля payload: {0}' -f ($stmt.FieldNames -join ','))

# --- маска ключевых слов и уровень ------------------------------------------
$kwUnknown = @($Keywords | Where-Object { -not $manifest.Keywords.ContainsKey($_) })
if ($kwUnknown.Count -gt 0) {
    Write-Host ('Неизвестные ключевые слова: {0}' -f ($kwUnknown -join ', ')) -ForegroundColor Red
    Write-Host ('Провайдер знает такие: {0}' -f (($manifest.Keywords.Keys | Sort-Object) -join ', '))
    Write-Log ('неизвестные ключевые слова: ' + ($kwUnknown -join ', '))
    $script:Log.Dispose()
    exit 5
}
# ETW отдаёт событие, если хотя бы один его бит есть в маске сессии, поэтому
# служебный старший бит принадлежности к каналу добавлять не нужно
$kwMask = [uint64]0
foreach ($n in $Keywords) { $kwMask = $kwMask -bor $manifest.Keywords[$n] }

# уровень тоже не зашит: берём самый подробный среди событий, попадающих в маску
# (Informational=4 у оператора C/AL, Error=2 у сбоя функции - фильтр "не выше N"
# пропускает оба, если N=4)
$level = 0
foreach ($i in $manifest.Events.Values) {
    if (($i.KeywordMask -band $kwMask) -ne 0 -and $i.LevelValue -gt $level) { $level = $i.LevelValue }
}
if ($level -lt 4) { $level = 4 }
Write-Step ('маска: 0x{0:X} ({1}), уровень {2}' -f $kwMask, ($Keywords -join '+'), $level)

# --- конфигурация сервера ---------------------------------------------------
$svcRoot = Resolve-NavServicePath -Instance $ServerInstance -Hint $NavServicePath
if (-not $svcRoot) {
    Write-Host 'Не найден каталог Service установки NAV. Укажите его через -NavServicePath.' -ForegroundColor Red
    $script:Log.Dispose()
    exit 6
}
$cfgPath = Resolve-NavConfigPath -ServiceRoot $svcRoot -Instance $ServerInstance
if (-not $cfgPath) {
    Write-Host ('Не найден CustomSettings.config в {0}.' -f $svcRoot) -ForegroundColor Red
    $script:Log.Dispose()
    exit 6
}
Write-Step ('конфигурация: {0}' -f $cfgPath)

# файл переписывается прицельно, регулярным выражением по одному атрибуту:
# XmlDocument.Save переформатировал бы весь файл в Program Files
$cfgBytesOrig = [System.IO.File]::ReadAllBytes($cfgPath)
$hasBom = ($cfgBytesOrig.Length -ge 3 -and $cfgBytesOrig[0] -eq 0xEF -and
           $cfgBytesOrig[1] -eq 0xBB -and $cfgBytesOrig[2] -eq 0xBF)
if ($hasBom) {
    $cfgTextOrig = [System.Text.Encoding]::UTF8.GetString($cfgBytesOrig, 3, $cfgBytesOrig.Length - 3)
} else {
    $cfgTextOrig = [System.Text.Encoding]::UTF8.GetString($cfgBytesOrig)
}
$rxSetting = New-Object System.Text.RegularExpressions.Regex(
    ('(<add\s+key\s*=\s*"{0}"\s+value\s*=\s*")([^"]*)(")' -f $SETTING_KEY), 'IgnoreCase')
$mSetting = $rxSetting.Match($cfgTextOrig)
if ($mSetting.Success) { $tracingOrig = $mSetting.Groups[2].Value } else { $tracingOrig = '<нет ключа>' }

$svcName = 'MicrosoftDynamicsNavServer$' + $ServerInstance
$svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
if ($svc) { $svcOrigStatus = [string]$svc.Status } else { $svcOrigStatus = '<нет службы>' }

# --- план изменений ---------------------------------------------------------
# печатается ДО первого изменяющего действия: видно и что меняем, и как вернём
$plan = New-Object System.Collections.Generic.List[object]
[void]$plan.Add([pscustomobject]@{
    What = ('Сессия ETW {0}' -f $SessionName); From = 'нет'
    To = ('в памяти, кольцо {0} МБ' -f $MaxSizeMB); Back = 'logman stop -ets' })
if ($tracingOrig -ne 'true') {
    [void]$plan.Add([pscustomobject]@{
        What = $SETTING_KEY; From = $tracingOrig; To = 'true'; Back = ('вернём ' + $tracingOrig) })
} else {
    [void]$plan.Add([pscustomobject]@{
        What = $SETTING_KEY; From = 'true'; To = 'true (не меняем)'; Back = 'нечего возвращать' })
}
if ($svc) {
    if ($NoRestart) { $to = 'без перезапуска (-NoRestart)' } else { $to = 'Running, с перезапуском' }
    [void]$plan.Add([pscustomobject]@{
        What = ('Служба ' + $svcName); From = $svcOrigStatus; To = $to; Back = ('вернём ' + $svcOrigStatus) })
}
if ($Calibrate) {
    [void]$plan.Add([pscustomobject]@{
        What = 'Калибровка'; From = 'нет'; To = 'сценарий выполняется ДВАЖДЫ'; Back = 'в базу ничего не пишется' })
}

Write-Host ''
Write-Host 'ПЛАН ИЗМЕНЕНИЙ' -ForegroundColor Cyan
Write-Host ('{0,-32} {1,-16} {2,-30} {3}' -f 'ЧТО', 'БЫЛО', 'СТАНЕТ', 'КАК ВЕРНЁМ')
Write-Host (('-' * 32) + ' ' + ('-' * 16) + ' ' + ('-' * 30) + ' ' + ('-' * 24))
foreach ($p in $plan) {
    Write-Host ('{0,-32} {1,-16} {2,-30} {3}' -f $p.What, $p.From, $p.To, $p.Back)
    Write-Log ('план: {0} | {1} -> {2} | откат: {3}' -f $p.What, $p.From, $p.To, $p.Back)
}
Write-Host ''
Write-Host 'Канал журнала Microsoft-DynamicsNAV-Server/Debug и постоянные настройки ETW НЕ трогаются.'
Write-Host ('Откат выполняется в блоке finally - и при ошибке, и при Ctrl+C.')
Write-Host ('Если процесс убьют: {0}' -f (Join-Path $OutDir 'restore.cmd'))
if ($KeepTracingOn) {
    Write-Host 'ВНИМАНИЕ: задан -KeepTracingOn - трассировка и служба останутся как есть.' -ForegroundColor Yellow
}

if (-not $Yes) {
    Write-Host ''
    $answer = Read-Host 'Продолжить? [y/N]'
    if ($answer -notmatch '^(y|yes|д|да)$') {
        Write-Host 'Отменено, ничего не изменено.' -ForegroundColor Yellow
        Write-Log 'отказ пользователя в плане изменений'
        $script:Log.Dispose()
        exit 9
    }
}

# --- аварийный откат на случай, если процесс убьют --------------------------
# .cmd пишется в кодировке OEM 866 - именно её ждёт cmd.exe
$cfgBak = Join-Path $OutDir 'CustomSettings.config.bak'
[System.IO.File]::WriteAllBytes($cfgBak, $cfgBytesOrig)
$restoreLines = @(
    '@echo off'
    'rem Аварийный откат настроек трассировки C/AL.'
    ('rem Создан {0:yyyy-MM-dd HH:mm:ss} скриптом Collect-AlTrace.ps1.' -f (Get-Date))
    'rem Запускать от имени администратора, если сбор был прерван жёстко.'
    ('logman stop "{0}" -ets' -f $SessionName)
    ('copy /Y "{0}" "{1}"' -f $cfgBak, $cfgPath)
)
if ($svc) {
    $restoreLines += ('net stop "{0}"' -f $svcName)
    if ($svcOrigStatus -eq 'Running') { $restoreLines += ('net start "{0}"' -f $svcName) }
}
$restoreLines += 'echo Настройки возвращены в исходное состояние.'
[System.IO.File]::WriteAllText((Join-Path $OutDir 'restore.cmd'),
    (($restoreLines -join "`r`n") + "`r`n"), [System.Text.Encoding]::GetEncoding(866))

# ---------------------------------------------------------------------------
# собственно сбор
# ---------------------------------------------------------------------------
$cfgChanged     = $false      # конфигурация сейчас в изменённом состоянии
$cfgEverChanged = $false      # конфигурацию трогали хоть раз за прогон
$svcTouched     = $false
$sessionRunning = $false
$scenarioSec    = 0.0
$baselineSec    = 0.0
$stats          = $null
$lossNote       = @()

function Set-TracingFlag {
    param([bool] $On)
    $value = 'false'
    if ($On) { $value = 'true' }
    if ($script:mSettingSuccess) {
        $newText = $script:rxSettingRef.Replace($script:cfgTextRef, ('${1}' + $value + '${3}'), 1)
    } else {
        $newText = $script:cfgTextRef -replace '(?i)</appSettings>',
            ('  <add key="{0}" value="{1}" />{2}</appSettings>' -f $SETTING_KEY, $value, "`r`n")
        if ($newText -eq $script:cfgTextRef) { throw 'В CustomSettings.config не найден раздел appSettings.' }
    }
    [System.IO.File]::WriteAllText($script:cfgPathRef, $newText,
        (New-Object System.Text.UTF8Encoding($script:cfgBomRef)))
    Write-Step ('{0} = {1}' -f $SETTING_KEY, $value)
}
$script:rxSettingRef    = $rxSetting
$script:cfgTextRef      = $cfgTextOrig
$script:cfgPathRef      = $cfgPath
$script:cfgBomRef       = $hasBom
$script:mSettingSuccess = $mSetting.Success

try {
    # ---- 0. калибровочный прогон БЕЗ трассировки --------------------------
    # цена события снимается разностью: тот же сценарий сначала без трассировки,
    # потом с ней. Никаких объектов в базе для этого не нужно.
    if ($Calibrate) {
        if ($tracingOrig -eq 'true') {
            Set-TracingFlag -On $false
            $cfgChanged = $true; $cfgEverChanged = $true
            if (-not $NoRestart -and $svc) { $svcTouched = $true; Set-NavServiceRunning -Service $svc -Restart }
        }
        if ($svc) { $svcTouched = $true; Set-NavServiceRunning -Service $svc }

        Write-Host ''
        Write-Host '================================================================' -ForegroundColor Green
        Write-Host ' ПРОГОН 1 из 2: БЕЗ трассировки (эталон)' -ForegroundColor Green
        Write-Host '================================================================' -ForegroundColor Green
        Write-Host ' Выполните сценарий в NAV. Второй прогон должен быть ТЕМ ЖЕ.'
        Write-Host ''
        $swBase = [System.Diagnostics.Stopwatch]::StartNew()
        Read-Host ' Нажмите Enter, когда сценарий отработает' | Out-Null
        $swBase.Stop()
        $baselineSec = $swBase.Elapsed.TotalSeconds
        Write-Step ('эталонный прогон: {0:N1} с' -f $baselineSec)
    }

    # ---- 1. трассировка в конфигурации ------------------------------------
    Write-Head 'Подготовка'
    $needRestart = $false
    if ($tracingOrig -ne 'true' -or $cfgChanged) {
        Set-TracingFlag -On $true
        $cfgChanged  = $true
        $cfgEverChanged = $true
        $needRestart = $true
    } else {
        Write-Step ('{0} уже true - конфигурация не менялась' -f $SETTING_KEY)
    }

    # ---- 2. служба: значение читается при старте ---------------------------
    if ($svc) {
        $svcTouched = $true
        if ($needRestart -and -not $NoRestart) { Set-NavServiceRunning -Service $svc -Restart }
        else                                   { Set-NavServiceRunning -Service $svc }
        $svc.Refresh()
        if ($needRestart -and $NoRestart) {
            Write-Host (' ВНИМАНИЕ: {0} изменён, но инстанция не перезапускалась -' -f $SETTING_KEY) -ForegroundColor Yellow
            Write-Host ' событий уровня оператора не будет.' -ForegroundColor Yellow
        }
    } else {
        Write-Step ('служба {0} не найдена - сессия соберёт только то, что пишут другие процессы' -f $svcName)
    }

    # ---- 3. эфемерная сессия ETW ------------------------------------------
    # хвост прошлого прогона, если он был убит: логично снять молча
    $leftover = Invoke-Native -Exe 'logman.exe' -Arguments @('stop', $SessionName, '-ets') -AllowFail
    if ($leftover.Code -eq 0) { Write-Step ('снята сессия {0}, оставшаяся от прошлого прогона' -f $SessionName) }
    if (Test-Path -LiteralPath $etlPath) { [System.IO.File]::Delete($etlPath) }

    $provGuid = '{' + $manifest.Provider.Id.ToString() + '}'
    $common = @('-o', $etlPath, '-mode', 'Circular', '-max', [string]$MaxSizeMB,
                '-bs', [string]$BufferSizeKB, '-nb', [string]$MinBuffers, [string]$MaxBuffers,
                '-ct', 'perf', '-ets')
    # провайдер задаётся именем; если платформа отдала его только по GUID -
    # вторая попытка идёт с GUID из манифеста, третья - в формате bincirc
    $attempts = @(
        @{ Note = 'по имени провайдера'; Args = (@('create', 'trace', $SessionName, '-p', $PROVIDER, ('0x{0:X}' -f $kwMask), [string]$level) + $common) }
        @{ Note = 'по GUID провайдера';  Args = (@('create', 'trace', $SessionName, '-p', $provGuid, ('0x{0:X}' -f $kwMask), [string]$level) + $common) }
        @{ Note = 'по GUID, формат bincirc'; Args = (@('create', 'trace', $SessionName, '-p', $provGuid, ('0x{0:X}' -f $kwMask), [string]$level,
                     '-o', $etlPath, '-f', 'bincirc', '-max', [string]$MaxSizeMB, '-bs', [string]$BufferSizeKB,
                     '-nb', [string]$MinBuffers, [string]$MaxBuffers, '-ct', 'perf', '-ets')) }
    )
    $created = $false
    foreach ($a in $attempts) {
        $r = Invoke-Native -Exe 'logman.exe' -Arguments $a.Args -AllowFail
        if ($r.Code -eq 0) {
            $created = $true
            Write-Step ('сессия {0} создана ({1})' -f $SessionName, $a.Note)
            break
        }
        Write-Step ('попытка "{0}" не прошла, код {1}' -f $a.Note, $r.Code)
        # неудачная попытка могла оставить полусозданную сессию: следующая тогда
        # упала бы с "имя занято", а причина осталась бы неочевидной
        Invoke-Native -Exe 'logman.exe' -Arguments @('stop', $SessionName, '-ets') -AllowFail | Out-Null
    }
    if (-not $created) {
        Write-Host ''
        Write-Host ('Не удалось создать сессию ETW {0}. Подробности в collect.log.' -f $SessionName) -ForegroundColor Red
        Write-Host 'Проверьте: консоль поднята от администратора; имя сессии не занято' -ForegroundColor Yellow
        Write-Host ('(logman query -ets); на диске есть место под {0} МБ.' -f $MaxSizeMB) -ForegroundColor Yellow
        throw ('сессия ETW не создана: ' + $SessionName)
    }
    $sessionRunning = $true
    $sessionStart   = Get-Date
    Write-Log ('сессия поднята в {0:HH:mm:ss.fff}' -f $sessionStart)

    # ---- 4. сценарий -------------------------------------------------------
    Write-Host ''
    Write-Host '================================================================' -ForegroundColor Green
    if ($Calibrate) { Write-Host ' ПРОГОН 2 из 2: ТРАССИРОВКА ВКЛЮЧЕНА' -ForegroundColor Green }
    else            { Write-Host ' ТРАССИРОВКА ВКЛЮЧЕНА' -ForegroundColor Green }
    Write-Host '================================================================' -ForegroundColor Green
    Write-Host ' Выполните сценарий в NAV. Чем короче сценарий, тем чище замер:'
    Write-Host ' событие пишется на КАЖДЫЙ оператор C/AL всей инстанции.'
    Write-Host ''
    $swRun = [System.Diagnostics.Stopwatch]::StartNew()
    Read-Host ' Нажмите Enter, когда сценарий отработает' | Out-Null
    $swRun.Stop()
    $scenarioSec = $swRun.Elapsed.TotalSeconds

    # ---- 5. счётчики сессии и останов --------------------------------------
    # счётчик потерь живёт в свойствах ЖИВОЙ сессии, поэтому снимается до stop.
    # Подписи logman локализованы, поэтому ищем и английские, и русские.
    Write-Head 'Остановка сессии'
    $q = Invoke-Native -Exe 'logman.exe' -Arguments @('query', $SessionName, '-ets') -AllowFail
    foreach ($ln in $q.Output) {
        if ($ln -match '(?i)^\s*([^:]*(lost|потер|утрач)[^:]*):\s*(\d+)') {
            $lossNote += ('{0} = {1}' -f $Matches[1].Trim(), $Matches[3])
        }
    }
    Invoke-Native -Exe 'logman.exe' -Arguments @('stop', $SessionName, '-ets') | Out-Null
    $sessionRunning = $false
    Write-Step 'сессия остановлена, файл .etl закрыт'

    if (-not (Test-Path -LiteralPath $etlPath)) { throw ('сессия не оставила файла ' + $etlPath) }

    # ---- 6. разбор ---------------------------------------------------------
    Write-Head 'Разбор событий'
    $stats = ConvertTo-TraceTsv -Manifest $manifest -TracePath $etlPath -TsvPath $tsvPath -ProviderName $PROVIDER
    foreach ($v in $stats.Verify) { Write-Log ('сверка полей: ' + $v) }

    # ---- 7. итог -----------------------------------------------------------
    Write-TraceSummary -Stats $stats -TracePath $etlPath -ScenarioSec $scenarioSec

    # Потери в кольцевой сессии выглядят иначе, чем в журнале: файл не переполняется,
    # а затирает НАЧАЛО трассы. Признак - первое событие заметно позже старта сессии
    # при файле, дотянувшем до предела.
    $etlMB   = (Get-Item -LiteralPath $etlPath).Length / 1MB
    $lagSec  = 0.0
    if ($stats.Total -gt 0 -and $stats.TickMin -lt [int64]::MaxValue) {
        $lagSec = (([datetime]$stats.TickMin) - $sessionStart).TotalSeconds
    }
    $wrapped = ($etlMB -ge ($MaxSizeMB * 0.98)) -and ($lagSec -gt 5.0)

    Write-Host ('Кольцо:    {0:N1} из {1} МБ; первое событие через {2:N1} с после старта сессии' -f
        $etlMB, $MaxSizeMB, $lagSec)
    if ($lossNote.Count -gt 0) {
        Write-Host ('Счётчики:  {0}' -f ($lossNote -join '; '))
    } else {
        Write-Host 'Счётчики:  logman не отдал счётчик потерь (см. collect.log)' -ForegroundColor DarkGray
    }

    if ($stats.Total -eq 0) {
        Write-Host 'Потеряно:  событий нет вовсе' -ForegroundColor Red
        Write-Host ('  Проверьте по порядку: {0}=true; инстанция ПЕРЕЗАПУЩЕНА после' -f $SETTING_KEY) -ForegroundColor Yellow
        Write-Host '  включения; сценарий действительно выполнял код C/AL; маска -Keywords' -ForegroundColor Yellow
        Write-Host '  не отсекла нужный класс событий.' -ForegroundColor Yellow
        $exitCode = 8
    }
    elseif ($wrapped) {
        Write-Host ('Потеряно:  кольцо перезаписало начало трассы - первые {0:N1} с сценария НЕ в файле' -f $lagSec) -ForegroundColor Red
        Write-Host ('  Повторите с -MaxSizeMB {0} или сократите сценарий.' -f ([math]::Min(16384, $MaxSizeMB * 4))) -ForegroundColor Yellow
        $exitCode = 8
    }
    elseif (($lossNote -join ' ') -match '=\s*[1-9]') {
        Write-Host 'Потеряно:  ядро сообщило о потерянных буферах - см. счётчики выше' -ForegroundColor Yellow
        Write-Host ('  Повторите с -BufferSizeKB {0} -MaxBuffers {1}.' -f ($BufferSizeKB * 2), ($MaxBuffers * 2)) -ForegroundColor Yellow
        # Единственная из трёх веток потерь, которая кода возврата НЕ выставляла, - и это
        # было упущение, а не решение: вызывающий трактует 8 как «трасса пустая ЛИБО
        # ПОТЕРЯНА». Терялись целые буферы, Hits и Self занижались неравномерно, а прогон
        # объявлялся зелёным.
        $exitCode = 8
    }
    else {
        Write-Host 'Потеряно:  0' -ForegroundColor Green
    }

    # ---- 8. цена события ---------------------------------------------------
    if ($Calibrate) {
        $delta = $scenarioSec - $baselineSec
        Write-Host ''
        Write-Host ('Калибровка: без трассировки {0:N1} с, с трассировкой {1:N1} с, разность {2:N1} с' -f
            $baselineSec, $scenarioSec, $delta)
        if ($stats.Total -gt 0 -and $delta -gt 0) {
            $us = ($delta * 1000000.0) / $stats.Total
            Write-Host ('Цена события: {0:N1} мкс на событие ({1:N0} событий)' -f $us, $stats.Total) -ForegroundColor Cyan
            Write-Log ('цена события: {0:N3} мкс' -f $us)
        } else {
            Write-Host 'Цена события: измерить не удалось (разность неположительна либо событий нет)' -ForegroundColor Yellow
        }
    }

    Write-Host ''
    Write-Host ('Файлы:     events.tsv, trace.etl, manifest.tsv, collect.log')
    Write-Host ('Каталог:   {0}' -f $OutDir)
    Write-Log ('итог Full: событий {0}, кольцо {1:N1} МБ, сценарий {2:N1} с' -f $stats.Total, $etlMB, $scenarioSec)
}
catch {
    Write-Host ''
    Write-Host ('ОШИБКА: {0}' -f $_.Exception.Message) -ForegroundColor Red
    Write-Log ('ОШИБКА: ' + $_.Exception.ToString())
    if ($exitCode -eq 0) { $exitCode = 1 }
}
finally {
    # Откат обязателен и при ошибке, и при Ctrl+C: иначе на машине останется живая
    # сессия ETW и включённая трассировка, которая сажает инстанцию до следующего
    # вмешательства. Порядок обратный порядку изменений: сначала сессия (пока она
    # жива, .etl не дочитать), потом конфигурация, потом служба.
    Write-Head 'Откат'
    if ($sessionRunning) {
        Invoke-Native -Exe 'logman.exe' -Arguments @('stop', $SessionName, '-ets') -AllowFail | Out-Null
        $sessionRunning = $false
        Write-Step ('сессия {0} остановлена' -f $SessionName)
    } else {
        Write-Step 'сессия ETW уже остановлена'
    }

    if ($KeepTracingOn) {
        Write-Host ''
        Write-Host 'ТРАССИРОВКА ОСТАВЛЕНА ВКЛЮЧЁННОЙ (-KeepTracingOn).' -ForegroundColor Yellow
        Write-Host ('Вернуть как было: {0}' -f (Join-Path $OutDir 'restore.cmd')) -ForegroundColor Yellow
        Write-Log 'откат конфигурации и службы пропущен по -KeepTracingOn'
    }
    else {
        if ($cfgChanged) {
            try {
                [System.IO.File]::WriteAllBytes($cfgPath, $cfgBytesOrig)
                $cfgChanged = $false
                Write-Step ('{0} возвращён в {1}' -f $SETTING_KEY, $tracingOrig)
            }
            catch {
                Write-Host ('НЕ УДАЛОСЬ вернуть CustomSettings.config: {0}' -f $_.Exception.Message) -ForegroundColor Red
                Write-Host ('Восстановите вручную: {0}' -f (Join-Path $OutDir 'restore.cmd')) -ForegroundColor Red
                Write-Log 'откат конфигурации НЕ УДАЛСЯ'
                # Машина осталась изменённой - это требует рук на сервере и перебивает
                # любой результат сбора. Раньше все присвоения кода лежали ДО finally,
                # и провал отката выходил с нулём: вызывающий печатал зелёное «ОК».
                $exitCode = 9
            }
        }
        if ($svcTouched -and $svc) {
            try {
                $svc.Refresh()
                if ($svcOrigStatus -eq 'Running') {
                    # конфигурацию вернули, но работающая инстанция всё ещё держит
                    # старое значение в памяти: без перезапуска трассировка осталась бы
                    # включённой. Если конфигурацию не трогали - и перезапускать нечего.
                    if ($cfgEverChanged -and -not $NoRestart) { Set-NavServiceRunning -Service $svc -Restart }
                    else                                      { Set-NavServiceRunning -Service $svc }
                } else {
                    Stop-NavService -Service $svc
                }
                Write-Step ('служба возвращена в состояние {0}' -f $svcOrigStatus)
            }
            catch {
                Write-Host ('НЕ УДАЛОСЬ вернуть службу в состояние {0}: {1}' -f $svcOrigStatus, $_.Exception.Message) -ForegroundColor Red
                Write-Log 'откат службы НЕ УДАЛСЯ'
                # То же самое: инстанция не поднялась - на сервере остались руки, а не
                # успешный прогон. Код 9 перебивает 8: разбираться надо сперва с машиной.
                $exitCode = 9
            }
        }
    }
    if ($script:Log) { $script:Log.Flush(); $script:Log.Dispose(); $script:Log = $null }
}

exit $exitCode
