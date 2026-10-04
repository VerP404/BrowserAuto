# Параллельный ввод талонов цели 3 (БСК / multi-service).
#
# Примеры (из корня BrowserAuto):
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Talons "...\талоны.xlsx" -Services "...\услуги.xlsx"
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5                  # все категории
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Category "BSK"  # только БСК

param(
    [int]$Workers = 5,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$Talons = "",
    [string]$Services = "",
    [string]$Building = "",
    [string]$Doctor = "",
    # Пусто = все категории. ASCII-алиас BSK = БСК.
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
$TalonsXlsx = if ($Talons) { (Resolve-Path $Talons).Path } else { Join-Path $DataDir "талоны.xlsx" }
$ServicesXlsx = if ($Services) { (Resolve-Path $Services).Path } else { Join-Path $DataDir "услуги.xlsx" }

if (-not (Test-Path $Py)) { throw "Нет venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "Нет лаунчера: $Launcher" }
if (-not (Test-Path $CountPy)) { throw "Нет счётчика: $CountPy" }
if (-not (Test-Path $TalonsXlsx)) { throw "Нет Excel талонов: $TalonsXlsx" }
if (-not (Test-Path $ServicesXlsx)) { throw "Нет Excel услуг: $ServicesXlsx" }

# UTF-8 для Python (кириллица в argv / stdout)
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# Нормализация алиасов категории
$CategoryNorm = $Category
if ($CategoryNorm -eq "BSK" -or $CategoryNorm -eq "bsk") { $CategoryNorm = "БСК" }
if ($CategoryNorm -eq "ONKO" -or $CategoryNorm -eq "onko") { $CategoryNorm = "ОНКО" }
if ($CategoryNorm -eq "SD" -or $CategoryNorm -eq "sd") { $CategoryNorm = "СД" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$total = & $Py $CountPy $TalonsXlsx $CategoryNorm
if ($LASTEXITCODE -ne 0) { throw "Не удалось посчитать строки талонов" }
$total = [int]$total
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Нечего обрабатывать: total=$total startOffset=$StartOffset category=$CategoryNorm"
    Write-Host "Подсказка: -Category '' (все) или -Category BSK (ASCII = БСК)"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk category=$CategoryNorm"
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
Write-Host "Запущено: $($procs.Count). Логи: $LogDir\w*.err.log"
Write-Host "CSV: $DataDir\oms_3_result_*.csv"
