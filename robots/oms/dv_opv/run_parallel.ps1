# Parallel DV4/OPV workers (Start-Process + RedirectStandard*).
# From BrowserAuto root:
#   .\robots\oms\dv_opv\run_parallel.ps1 -Workers 6
#   .\robots\oms\dv_opv\run_parallel.ps1 -Workers 6 -File "...\talons_dv_opv_doc_30151001.xlsx"

param(
    [int]$Workers = 6,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$File = "",
    [string]$Catalog = "",
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_medical_exam.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"
$DefaultFile = Join-Path $DataDir "talons_dv_opv_doc_30151001.xlsx"

if (-not $Catalog) {
    # Prefer non-catalog xlsx in robot folder (full dictionary), else catalog_dv_opv.xlsx
    $xlsxHere = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter "*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '^~\$' })
    $dict = @($xlsxHere | Where-Object { $_.Name -notmatch '^catalog_' } | Sort-Object Length -Descending)
    if ($dict.Count -ge 1) {
        $Catalog = $dict[0].FullName
    } elseif (Test-Path -LiteralPath (Join-Path $PSScriptRoot "catalog_dv_opv.xlsx")) {
        $Catalog = Join-Path $PSScriptRoot "catalog_dv_opv.xlsx"
    } elseif ($xlsxHere.Count -ge 1) {
        $Catalog = $xlsxHere[0].FullName
    }
}

$Xlsx = if ($File) { (Resolve-Path -LiteralPath $File).Path } else { (Resolve-Path -LiteralPath $DefaultFile).Path }
if (-not $Catalog -or -not (Test-Path -LiteralPath $Catalog)) {
    throw "Catalog xlsx not found in $PSScriptRoot (pass -Catalog)"
}
$Cat = (Resolve-Path -LiteralPath $Catalog).Path

if (-not (Test-Path $Py)) { throw "No venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "No launcher: $Launcher" }
if (-not (Test-Path -LiteralPath $Xlsx)) { throw "No Excel: $Xlsx" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$tmpPy = Join-Path $env:TEMP "dv_opv_count_rows.py"
@(
    "import sys"
    "from openpyxl import load_workbook"
    "wb = load_workbook(sys.argv[1], read_only=True, data_only=True)"
    "ws = wb.active"
    "n = 0"
    "it = ws.iter_rows(values_only=True)"
    "next(it, None)"
    "for r in it:"
    "    if not r or all(v is None or str(v).strip()=='' for v in r):"
    "        continue"
    "    n += 1"
    "print(n)"
    "wb.close()"
) | Set-Content -Path $tmpPy -Encoding ASCII

$total = [int]((& $Py $tmpPy $Xlsx) | Select-Object -Last 1).ToString().Trim()
Remove-Item $tmpPy -Force -ErrorAction SilentlyContinue

$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Nothing to do: total=$total startOffset=$StartOffset"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk"
Write-Host "file=$Xlsx"
Write-Host "catalog=$Cat"

$procs = @()
for ($w = 0; $w -lt $Workers; $w++) {
    $offset = $StartOffset + ($w * $chunk)
    if ($offset -ge $total) { break }

    $limit = [Math]::Min($chunk, $total - $offset)
    if ($LimitPerWorker -gt 0) {
        $limit = [Math]::Min($limit, $LimitPerWorker)
    }
    if ($limit -le 0) { continue }

    $worker = "opv_w$w"
    $outLog = Join-Path $LogDir "$worker.out.log"
    $errLog = Join-Path $LogDir "$worker.err.log"
    Remove-Item $outLog, $errLog -Force -ErrorAction SilentlyContinue

    $argList = @(
        $Launcher,
        "--file", $Xlsx,
        "--catalog", $Cat,
        "--offset", "$offset",
        "--limit", "$limit"
    )

    Write-Host "Start $worker offset=$offset limit=$limit"

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
        OutLog = $outLog
        ErrLog = $errLog
    }
    Start-Sleep -Milliseconds $StaggerMs
}

$procs | Format-Table -AutoSize
$procs | ConvertTo-Json | Set-Content (Join-Path $LogDir "opv_pids.json") -Encoding utf8

Write-Host ""
Write-Host "Started workers: $($procs.Count)"
Write-Host "Logs: $LogDir\opv_w*.err.log"
Write-Host "CSV:  $DataDir\dv_opv_result_*.csv"
Write-Host "Watch: Get-Content $LogDir\opv_w0.err.log -Wait -Tail 30"
Write-Host "PIDs:  $LogDir\opv_pids.json"
