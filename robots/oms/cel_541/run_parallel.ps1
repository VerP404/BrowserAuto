# Параллельный ввод талонов 541.
#   .\robots\oms\cel_541\run_parallel.ps1 -Workers 5 -Building "ГП №3"
#   .\robots\oms\cel_541\run_parallel.ps1 -Workers 4 -StartOffset 0 -LimitPerWorker 10

param(
    [int]$Workers = 5,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$Talons = "",
    [string]$Services = "",
    [string]$Building = "",
    [string]$Doctor = "",
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_541.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"
$TalonsXlsx = if ($Talons) { (Resolve-Path $Talons).Path } else { Join-Path $DataDir "talons_541.xlsx" }
$ServicesXlsx = if ($Services) { (Resolve-Path $Services).Path } else { Join-Path $DataDir "services_541.xlsx" }

if (-not (Test-Path $Py)) { throw "Нет venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "Нет лаунчера: $Launcher" }
if (-not (Test-Path $TalonsXlsx)) { throw "Нет Excel талонов: $TalonsXlsx" }
if (-not (Test-Path $ServicesXlsx)) { throw "Нет Excel услуг: $ServicesXlsx" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$talonsPy = $TalonsXlsx.Replace("\", "\\")
$total = & $Py -c "from openpyxl import load_workbook; wb=load_workbook(r'$talonsPy', read_only=True); print(wb.active.max_row-1); wb.close()"
$total = [int]$total
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Нечего обрабатывать: total=$total startOffset=$StartOffset"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk"
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
        "--window-y", "$y"
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
Write-Host "Запущено: $($procs.Count). Логи: $LogDir\w*.err.log"
Write-Host "CSV: $DataDir\oms_541_result_*.csv"
