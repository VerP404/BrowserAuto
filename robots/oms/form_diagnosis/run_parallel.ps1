# Параллельное исправление диагноза (form_diagnosis).
# Примеры (из корня BrowserAuto):
#   .\robots\oms\form_diagnosis\run_parallel.ps1 -Workers 5 -File "D:\path\Книга15.xlsx"
#   .\robots\oms\form_diagnosis\run_parallel.ps1 -Workers 5 -File "...\Книга15.xlsx" -Diagnosis "Z72.4"

param(
    [int]$Workers = 5,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$File = "",
    [string]$Diagnosis = "",
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_form_diagnosis.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"
$Xlsx = if ($File) { (Resolve-Path $File).Path } else { Join-Path $DataDir "Книга15.xlsx" }

if (-not (Test-Path $Py)) { throw "Нет venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "Нет лаунчера: $Launcher" }
if (-not (Test-Path $Xlsx)) { throw "Нет Excel: $Xlsx" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

$xlsxPy = $Xlsx.Replace("\", "\\")
$total = & $Py -c "from openpyxl import load_workbook; wb=load_workbook(r'$xlsxPy', read_only=True); ws=wb.active; print(max(ws.max_row-1,0)); wb.close()"
$total = [int]$total
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Нечего обрабатывать: total=$total startOffset=$StartOffset"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk"
Write-Host "file=$Xlsx diagnosis=$Diagnosis"

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
        "--file", $Xlsx,
        "--worker", $worker,
        "--offset", "$offset",
        "--limit", "$limit",
        "--window-x", "$x",
        "--window-y", "$y"
    )
    if ($Diagnosis) {
        $argList += @("--diagnosis", $Diagnosis)
    }

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
        ErrLog = $errLog
    }
    Start-Sleep -Milliseconds $StaggerMs
}

$procs | Format-Table -AutoSize
$procs | ConvertTo-Json | Set-Content (Join-Path $LogDir "pids.json") -Encoding utf8
Write-Host "Запущено: $($procs.Count). Логи: $LogDir\w*.err.log"
