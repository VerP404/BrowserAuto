# Параллельный ввод талонов 307 (5 воркеров по умолчанию).
# Пример:
#   .\robots\oms\cel_307\run_parallel.ps1 -Workers 5 -StartOffset 5
#   .\robots\oms\cel_307\run_parallel.ps1 -Workers 5 -StartOffset 10

param(
    [int]$Workers = 5,
    [int]$StartOffset = 5,
    [int]$LimitPerWorker = 0,
    [string]$File = "",
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_307.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"
$Xlsx = if ($File) { $File } else { Join-Path $DataDir "talons_307.xlsx" }

if (-not (Test-Path $Py)) { throw "Нет venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "Нет лаунчера: $Launcher" }
if (-not (Test-Path $Xlsx)) { throw "Нет Excel: $Xlsx" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$total = & $Py -c "from openpyxl import load_workbook; wb=load_workbook(r'$Xlsx', read_only=True); print(wb.active.max_row-1); wb.close()"
$total = [int]$total
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Нечего обрабатывать: total=$total startOffset=$StartOffset"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)

Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk"

$procs = @()
for ($w = 0; $w -lt $Workers; $w++) {
    $offset = $StartOffset + ($w * $chunk)
    if ($offset -ge $total) { break }
    $limit = [Math]::Min($chunk, $total - $offset)
    if ($LimitPerWorker -gt 0) {
        $limit = [Math]::Min($limit, $LimitPerWorker)
    }
    $x = 10 + $w * 180
    $y = 10 + ($w % 2) * 60
    $log = Join-Path $LogDir "w$w.log"
    Remove-Item $log -Force -ErrorAction SilentlyContinue

    # cmd /c: stdout+stderr в один файл (loguru пишет в stderr)
    $cmd = @(
        "`"$Py`"",
        "`"$Launcher`"",
        "--file", "`"$Xlsx`"",
        "--worker", "w$w",
        "--offset", "$offset",
        "--limit", "$limit",
        "--window-x", "$x",
        "--window-y", "$y",
        "1>`"$log`"",
        "2>&1"
    ) -join " "

    Write-Host "Start w$w offset=$offset limit=$limit window=($x,$y) log=$log"
    $p = Start-Process -FilePath "cmd.exe" -ArgumentList @("/c", $cmd) `
        -WorkingDirectory $Root `
        -WindowStyle Hidden `
        -PassThru
    $procs += [pscustomobject]@{
        Worker = "w$w"; Offset = $offset; Limit = $limit; Pid = $p.Id; Log = $log
    }
    Start-Sleep -Milliseconds $StaggerMs
}

$procs | Format-Table -AutoSize
$procs | ConvertTo-Json | Set-Content (Join-Path $LogDir "pids.json") -Encoding utf8
Write-Host "Логи: $LogDir\w*.log"
Write-Host "CSV:  $DataDir\oms_307_result_*.csv"
Write-Host "Следить: Get-Content robots\oms\cel_307\data\logs\w0.log -Wait -Tail 20"
