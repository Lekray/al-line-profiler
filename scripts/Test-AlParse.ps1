#Requires -Version 5.1
<#
.SYNOPSIS
    Прогон ядра статического разбора C/AL по штатным объектам NAV.

.DESCRIPTION
    Берёт дамп исходников (.alsrc, снимается Dump-AlSource.ps1) и гоняет по нему всю
    цепочку разбора: листинг, лексер, структуру функций и циклов, таблицу символов,
    цепочки фильтров. По каждому объекту печатает, сколько нашлось функций, циклов,
    переменных-записей и цепочек, сколько функций разобрано ненадёжно и за сколько
    миллисекунд это сделано.

    Отдельная проверка — области символов, которым в дампе нет функции: это признак
    того, что дамп устарел относительно скомпилированного объекта.

    Объекты взяты штатные — они есть в любой базе. Имена функций не задаются: состав
    функций меняется от версии к версии, поэтому для ручной сверки циклов берётся
    самая петлистая функция объекта.

.EXAMPLE
    powershell -File scripts\Test-AlParse.ps1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$scripts = $PSScriptRoot
. (Join-Path $scripts 'Lib-AlListing.ps1')
. (Join-Path $scripts 'Lib-AlParse.ps1')

$objects = @(
    @(5, 80), @(5, 90), @(5, 12), @(5, 22),
    @(1, 18), @(1, 36), @(1, 37), @(1, 38), @(1, 39)
)

$rows = @()
$allStruct = @{}
foreach ($o in $objects) {
    $t = $o[0]; $id = $o[1]
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $listing = ConvertTo-AlRows (Get-AlListing -ObjectType $t -ObjectId $id)
    $lexed   = ConvertTo-AlRows (Get-AlLexed -Listing $listing)
    $symbols = ConvertTo-AlRows (Get-AlSymbols -ObjectType $t -ObjectId $id)
    $struct  = Get-AlStructure -Listing $listing -Lexed $lexed
    $chains  = ConvertTo-AlRows (Get-AlFilterChains -Listing $listing -Lexed $lexed -Structure $struct -Symbols $symbols -ObjectType $t -ObjectId $id)
    $sw.Stop()
    $allStruct["$t`_$id"] = @{ S = $struct; L = $listing; Sym = $symbols; C = $chains }

    $funcs   = @($struct.Functions)
    $bad     = @($funcs | Where-Object { -not $_.Reliable })
    $loops   = 0; foreach ($f in $funcs) { $loops += $f.LoopCount }
    $recVars = @($symbols | Where-Object { $_.Kind -eq 'Record' })
    $recMapped = @($recVars | Where-Object { $_.ObjectId -gt 0 })

    # сверка областей символов с картой функций дампа (сигнал устаревания baseline)
    $dumpNames = @{}
    foreach ($f in $funcs) { $dumpNames[$f.Name] = $true }
    $scopes = @($symbols | Where-Object { $_.Scope } | ForEach-Object { $_.Scope } | Sort-Object -Unique)
    $orphan = @($scopes | Where-Object { -not $dumpNames.ContainsKey($_) })

    $rows += [pscustomobject]@{
        'Объект'      = ('{0} {1}' -f (Get-AlTypeName $t), $id)
        'Функций'     = $funcs.Count
        'Ненадёжных'  = $bad.Count
        'Циклов'      = $loops
        'RecVars'     = $recVars.Count
        'RecСТабл'    = $recMapped.Count
        'Областей'    = $scopes.Count
        'МимоДампа'   = $orphan.Count
        'Цепочек'     = $chains.Count
        'мс'          = [int]$sw.Elapsed.TotalMilliseconds
    }
    if ($bad.Count -gt 0) {
        Write-Host ("  !! {0} {1} — ненадёжные функции:" -f (Get-AlTypeName $t), $id) -ForegroundColor Yellow
        $bad | Select-Object -First 5 | ForEach-Object {
            Write-Host ("     {0} (стр. {1}): {2}" -f $_.Name, $_.HeaderLine, $_.Reason) -ForegroundColor DarkYellow
        }
    }
    if ($orphan.Count -gt 0) {
        Write-Host ("  !! {0} {1} — области символов без функции в дампе: {2}" -f `
            (Get-AlTypeName $t), $id, (($orphan | Select-Object -First 5) -join ', ')) -ForegroundColor Yellow
    }
}

$rows | Format-Table -AutoSize

# --- ручная сверка циклов ---------------------------------------------------
Write-Host "=== Циклы для ручной сверки ===" -ForegroundColor Cyan
# Имя функции не задаём: у штатного объекта состав функций меняется от версии
# к версии. Берём самую петлистую — её и сверять глазами по дампу.
$s90 = $allStruct['5_90'].S
$f90 = @($s90.Functions | Sort-Object LoopCount -Descending) | Select-Object -First 1
if ($f90) {
    Write-Host ("CU 90 {0} ({1}..{2}): циклов {3}" -f $f90.Name, $f90.FirstLine, $f90.LastLine, $f90.LoopCount)
    $f90.Loops | Format-Table Kind, StartLine, EndLine, Depth -AutoSize
}

$s80 = $allStruct['5_80'].S
$f80 = @($s80.Functions | Sort-Object LoopCount -Descending) | Select-Object -First 1
if ($f80) {
    Write-Host ("CU 80 {0}: циклов {1}" -f $f80.Name, $f80.LoopCount)
    $f80.Loops | Select-Object -First 6 | Format-Table Kind, StartLine, EndLine, Depth -AutoSize
}

$s37 = $allStruct['1_37'].S
$f37 = @($s37.Functions | Where-Object { $_.LoopCount -gt 0 } | Select-Object -First 2)
foreach ($ff in $f37) {
    Write-Host ("T37 {0}:" -f $ff.Name)
    $ff.Loops | Format-Table Kind, StartLine, EndLine, Depth -AutoSize
}

# --- пример цепочек фильтров ------------------------------------------------
Write-Host "=== Цепочки фильтров (CU 80, первые 6 с фильтрами) ===" -ForegroundColor Cyan
$allStruct['5_80'].C | Where-Object { $_.FilterLines.Count -gt 0 } | Select-Object -First 6 |
    ForEach-Object {
        Write-Host ("{0}: {1}.{2} стр.{3}  <- фильтры: {4}" -f $_.Function, $_.Variable, $_.ConsumerOp,
            $_.ConsumerLine, (($_.FilterLines | ForEach-Object { "$($_.Op)@$($_.LineNo)" }) -join ' '))
    }

# --- пример символов --------------------------------------------------------
Write-Host "=== Символы CU 80 (Record, первые 8) ===" -ForegroundColor Cyan
$allStruct['5_80'].Sym | Where-Object { $_.Kind -eq 'Record' } | Select-Object -First 8 |
    Format-Table Name, Scope, Origin, ObjectId, TableName, Temporary -AutoSize
