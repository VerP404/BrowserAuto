# Parallel sync of ALL services on ambulatory talons from portal zip.
#   .\robots\oms\fix_service\run_sync_parallel.ps1 -Workers 4 -Zip "D:\Projects\BrowserAuto\robots\oms\dn_fix_3_….zip"

param(
    [int]$Workers = 4,
    [string]$Zip = "",
    [int]$StaggerMs = 4000,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$Root = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $Root ".venv\Scripts\python.exe"
$Launcher = Join-Path $Root "run_sync_services.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs_sync"

if (-not $Zip) { throw "Pass -Zip path to portal bundle" }
$ZipPath = (Resolve-Path -LiteralPath $Zip).Path
if (-not (Test-Path $Py)) { throw "No venv: $Py" }

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

# count unique talons via sync_from_bundle
$countPy = Join-Path $PSScriptRoot "sync_from_bundle.py"
$countOut = & $Py -c @"
import sys
from pathlib import Path
sys.path.insert(0, r'$PSScriptRoot')
from sync_from_bundle import extract_bundle, load_bundle, RESULTS_DIR
z = Path(r'$ZipPath')
t, s = extract_bundle(z, RESULTS_DIR / '_sync_extract')
print(len(load_bundle(talons_path=t, services_path=s)))
"@
$total = [int](($countOut | Select-Object -Last 1).ToString().Trim())
if ($total -le 0) { throw "No talons in zip services" }

$chunk = [int][Math]::Ceiling($total / [double]$Workers)
Write-Host "total=$total workers=$Workers chunk=$chunk zip=$ZipPath"

function Quote-WinArg([string]$s) {
    if ($null -eq $s) { return '""' }
    if ($s -notmatch '[\s"]') { return $s }
    return '"' + ($s -replace '"', '\"') + '"'
}

$procs = @()
for ($w = 0; $w -lt $Workers; $w++) {
    $offset = $w * $chunk
    if ($offset -ge $total) { break }
    $limit = [Math]::Min($chunk, $total - $offset)
    $outLog = Join-Path $LogDir "sync_w$w.out.log"
    $errLog = Join-Path $LogDir "sync_w$w.err.log"
    Remove-Item $outLog, $errLog -Force -ErrorAction SilentlyContinue

    $args = @(
        $Launcher, "--zip", $ZipPath, "--offset", "$offset", "--limit", "$limit", "--worker", "w$w"
    )
    if ($DryRun) { $args += "--dry-run" }
    $argLine = ($args | ForEach-Object { Quote-WinArg $_ }) -join ' '
    Write-Host "Start w$w offset=$offset limit=$limit"
    $p = Start-Process -FilePath $Py -ArgumentList $argLine -WorkingDirectory $Root `
        -RedirectStandardOutput $outLog -RedirectStandardError $errLog -WindowStyle Hidden -PassThru
    $procs += [pscustomobject]@{ Worker = "w$w"; Offset = $offset; Limit = $limit; Pid = $p.Id }
    Start-Sleep -Milliseconds $StaggerMs
}

$procs | Format-Table -AutoSize
$procs | ConvertTo-Json | Set-Content (Join-Path $LogDir "sync_pids.json") -Encoding utf8
Write-Host "Logs: $LogDir\sync_w*.err.log"
