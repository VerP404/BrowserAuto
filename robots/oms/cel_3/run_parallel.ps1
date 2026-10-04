# Parallel cel_3 launcher (multi-service ambulatory).
# Use ASCII-only messages so PS5 on RU Windows does not break on encoding.
#
# Examples (from BrowserAuto root):
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Talons "D:\path\talony.xlsx" -Services "D:\path\uslugi.xlsx"
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Category BSK

param(
    [int]$Workers = 5,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$Talons = "",
    [string]$Services = "",
    [string]$Building = "",
    [string]$Doctor = "",
    # Empty = all categories. ASCII alias: BSK / ONKO / SD
    [string]$Category = "",
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_3.py"
$CountPy = Join-Path $PSScriptRoot "_count_talons.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"

if ($Talons) {
    $TalonsXlsx = (Resolve-Path -LiteralPath $Talons).Path
} else {
    $TalonsXlsx = Join-Path $DataDir ([string][char]0x0442 + [char]0x0430 + [char]0x043B + [char]0x043E + [char]0x043D + [char]0x044B + ".xlsx")
}
if ($Services) {
    $ServicesXlsx = (Resolve-Path -LiteralPath $Services).Path
} else {
    $ServicesXlsx = Join-Path $DataDir ([string][char]0x0443 + [char]0x0441 + [char]0x043B + [char]0x0443 + [char]0x0433 + [char]0x0438 + ".xlsx")
}

if (-not (Test-Path $Py)) { throw "venv python not found: $Py" }
if (-not (Test-Path $Launcher)) { throw "launcher not found: $Launcher" }
if (-not (Test-Path $CountPy)) { throw "counter not found: $CountPy" }
if (-not (Test-Path -LiteralPath $TalonsXlsx)) { throw "talons xlsx not found: $TalonsXlsx" }
if (-not (Test-Path -LiteralPath $ServicesXlsx)) { throw "services xlsx not found: $ServicesXlsx" }

$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# Category aliases -> Russian via [char] codes (no Cyrillic literals in .ps1)
$catKey = if ($null -eq $Category) { "" } else { $Category.Trim().ToLowerInvariant() }
$CategoryNorm = $Category
if ($catKey -eq "") {
    $CategoryNorm = ""
} elseif ($catKey -eq "bsk") {
    $CategoryNorm = ([string][char]0x0411 + [char]0x0421 + [char]0x041A)
} elseif ($catKey -eq "onko") {
    $CategoryNorm = ([string][char]0x041E + [char]0x041D + [char]0x041A + [char]0x041E)
} elseif ($catKey -eq "sd") {
    $CategoryNorm = ([string][char]0x0421 + [char]0x0414)
}

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$total = & $Py $CountPy $TalonsXlsx $CategoryNorm
if ($LASTEXITCODE -ne 0) { throw "failed to count talon rows" }
$total = [int]$total
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "nothing to process: total=$total startOffset=$StartOffset category='$CategoryNorm'"
    Write-Host "hint: omit -Category for all rows, or use -Category BSK"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk category='$CategoryNorm'"
Write-Host "talons=$TalonsXlsx"
Write-Host "services=$ServicesXlsx"

$procs = @()
for ($w = 0; $w -lt $Workers; $w++) {
    $offset = $StartOffset + ($w * $chunk)
    if ($offset -ge $total) { break }

    $limit = [Math]::Min($chunk, $total - $offset)
    if ($LimitPerWorker -gt 0) {
        $limit = [Math]::Min($limit, $LimitPerWorker)
    }
    if ($limit -le 0) { continue }

    $x = 10 + $w * 180
    $y = 10 + ($w % 2) * 60
    $worker = "w$w"
    $outLog = Join-Path $LogDir "$worker.out.log"
    $errLog = Join-Path $LogDir "$worker.err.log"
    Remove-Item $outLog, $errLog -Force -ErrorAction SilentlyContinue

    $argList = @(
        $Launcher,
        "--talons", $TalonsXlsx,
        "--services", $ServicesXlsx,
        "--worker", $worker,
        "--offset", "$offset",
        "--limit", "$limit",
        "--window-x", "$x",
        "--window-y", "$y",
        "--category", $CategoryNorm
    )
    if ($Building) { $argList += @("--building", $Building) }
    if ($Doctor) { $argList += @("--doctor", $Doctor) }

    Write-Host "Start $worker offset=$offset limit=$limit window=($x,$y)"

    $p = Start-Process -FilePath $Py `
        -ArgumentList $argList `
        -WorkingDirectory $Root `
        -RedirectStandardOutput $outLog `
        -RedirectStandardError $errLog `
        -WindowStyle Hidden `
        -PassThru

    $procs += [pscustomobject]@{
        Worker = $worker
        Offset = $offset
        Limit  = $limit
        Pid    = $p.Id
        ErrLog = $errLog
    }
    Start-Sleep -Milliseconds $StaggerMs
}

$procs | Format-Table -AutoSize
$procs | ConvertTo-Json | Set-Content (Join-Path $LogDir "pids.json") -Encoding utf8
Write-Host "started: $($procs.Count). logs: $LogDir\w*.err.log"
Write-Host "csv: $DataDir\oms_3_result_*.csv"
