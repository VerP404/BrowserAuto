# Параллельный ввод талонов цели 3 (БСК / multi-service).
#
# Примеры (из корня BrowserAuto):
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Talons "...\талоны.xlsx" -Services "...\услуги.xlsx"
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Category ""   # все категории
#   .\robots\oms\cel_3\run_parallel.ps1 -Workers 5 -Category "БСК" -LimitPerWorker 10

param(
    [int]$Workers = 5,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$Talons = "",
    [string]$Services = "",
    [string]$Building = "",
    [string]$Doctor = "",
    [string]$Category = "БСК",
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_3.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"
$TalonsXlsx = if ($Talons) { (Resolve-Path $Talons).Path } else { Join-Path $DataDir "талоны.xlsx" }
$ServicesXlsx = if ($Services) { (Resolve-Path $Services).Path } else { Join-Path $DataDir "услуги.xlsx" }

if (-not (Test-Path $Py)) { throw "Нет venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "Нет лаунчера: $Launcher" }
if (-not (Test-Path $TalonsXlsx)) { throw "Нет Excel талонов: $TalonsXlsx" }
if (-not (Test-Path $ServicesXlsx)) { throw "Нет Excel услуг: $ServicesXlsx" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

# Считаем строки ПОСЛЕ фильтра Категория (как в add_3.load_talons_bundle)
$talonsPy = $TalonsXlsx.Replace("\", "\\")
$catPy = $Category.Replace("\", "\\").Replace("'", "''")
$total = & $Py -c @"
from openpyxl import load_workbook
wb = load_workbook(r'$talonsPy', read_only=True, data_only=True)
ws = wb.active
rows = ws.iter_rows(values_only=True)
headers = [str(h or '').strip().lower() for h in next(rows)]
i_enp = next((i for i,h in enumerate(headers) if h in ('енп','enp')), None)
i_cat = next((i for i,h in enumerate(headers) if h == 'категория'), None)
cat = '$catPy'.strip().lower()
n = 0
for raw in rows:
    if not raw or i_enp is None:
        continue
    enp = raw[i_enp]
    if enp is None or str(enp).strip() == '':
        continue
    if cat and i_cat is not None:
        c = str(raw[i_cat] or '').strip().lower()
        if c != cat:
            continue
    n += 1
wb.close()
print(n)
"@
$total = [int]$total
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Нечего обрабатывать: total=$total startOffset=$StartOffset category=$Category"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk category=$Category"
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
    if ($null -ne $Category) { $argList += @("--category", $Category) }
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
