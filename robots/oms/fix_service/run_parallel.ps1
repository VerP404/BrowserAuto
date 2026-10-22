# Параллельная замена услуг (fix_service).
# Из корня BrowserAuto:
#   .\robots\oms\fix_service\run_parallel.ps1 -Workers 4 -File "...\fix_service.xlsx"
#   .\robots\oms\fix_service\run_parallel.ps1 -Workers 4 -File "...\talons.xlsx" -From B04.015.003 -To B04.047.003

param(
    [int]$Workers = 4,
    [int]$StartOffset = 0,
    [int]$LimitPerWorker = 0,
    [string]$File = "",
    [string]$From = "",
    [string]$To = "",
    [switch]$ReplaceAll,
    [int]$StaggerMs = 4000
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_fix_service.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"
$Xlsx = if ($File) { (Resolve-Path -LiteralPath $File).Path } else { Join-Path $DataDir "fix_service.xlsx" }

if (-not (Test-Path $Py)) { throw "Нет venv: $Py" }
if (-not (Test-Path $Launcher)) { throw "Нет лаунчера: $Launcher" }
if (-not (Test-Path -LiteralPath $Xlsx)) { throw "Нет Excel: $Xlsx" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

# Count via temp .py (PowerShell here-string mangles Cyrillic headers)
$tmpCount = Join-Path $env:TEMP "fix_service_count_rows.py"
@(
    "from openpyxl import load_workbook"
    "import sys"
    "wb=load_workbook(sys.argv[1], data_only=True)"
    "ws=wb.active"
    "headers=[(str(c.value).strip().lower() if c.value is not None else '') for c in next(ws.iter_rows(min_row=1,max_row=1))]"
    "names=('талон','номер','id','claim','n','web-oms','web_oms')"
    "i_t=next((i for i,h in enumerate(headers) if h in names), 0)"
    "n=0"
    "for r in ws.iter_rows(min_row=2, values_only=True):"
    "    if not r: continue"
    "    t=r[i_t] if i_t < len(r) else None"
    "    if t is not None and str(t).strip():"
    "        n+=1"
    "print(n)"
    "wb.close()"
) | Set-Content -Path $tmpCount -Encoding UTF8
$total = [int]((& $Py $tmpCount $Xlsx | Select-Object -Last 1).ToString().Trim())
Remove-Item $tmpCount -Force -ErrorAction SilentlyContinue
$remaining = [Math]::Max($total - $StartOffset, 0)
if ($remaining -le 0) {
    Write-Host "Нечего обрабатывать: total=$total startOffset=$StartOffset"
    exit 0
}

$chunk = [int][Math]::Ceiling($remaining / [double]$Workers)
Write-Host "total=$total startOffset=$StartOffset remaining=$remaining workers=$Workers chunk=$chunk"
Write-Host "file=$Xlsx from=$From to=$To"

function Quote-WinArg([string]$s) {
    if ($null -eq $s) { return '""' }
    if ($s -notmatch '[\s"]') { return $s }
    return '"' + ($s -replace '"', '\"') + '"'
}

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
    if ($From) { $argList += @("--from", $From) }
    if ($To) { $argList += @("--to", $To) }
    if ($ReplaceAll) { $argList += "--replace-all" }

    $argLine = ($argList | ForEach-Object { Quote-WinArg $_ }) -join " "
    Write-Host "Start $worker offset=$offset limit=$limit"

    $p = Start-Process -FilePath $Py `
        -ArgumentList $argLine `
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
