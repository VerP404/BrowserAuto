# Параллельный запуск add_carryover_plan через ASCII-лаунчер.
# Пример:
#   .\run_parallel.ps1 -Workers 6 -Limit 10
#   .\run_parallel.ps1 -Workers 6 -Limit 7000
# Полный прогон: задайте -Limit 0 и разные -ChunkSize (см. ниже).

param(
    [int]$Workers = 6,
    [int]$Limit = 10,
    [int]$ChunkSize = 0,
    [string]$File = "",
    [int]$Month = 0
)

$ErrorActionPreference = "Stop"
$BrowserAutoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$Py = Join-Path $BrowserAutoRoot ".venv\Scripts\python.exe"
$Launcher = Join-Path $BrowserAutoRoot "run_carryover_plan.py"
$DataDir = Join-Path $PSScriptRoot "data"
$LogDir = Join-Path $DataDir "logs"

if (-not (Test-Path $Py)) {
    throw "Не найден venv: $Py — создайте .venv и поставьте requirements.txt"
}
if (-not (Test-Path $Launcher)) {
    throw "Не найден лаунчер: $Launcher"
}

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

# Если ChunkSize не задан — режем по Limit на воркера (как тест 6×10)
if ($ChunkSize -le 0) {
    $ChunkSize = [Math]::Max($Limit, 1)
}

$procs = @()
for ($w = 0; $w -lt $Workers; $w++) {
    $offset = $w * $ChunkSize
    $x = 20 + $w * 160
    $y = 20 + ($w % 3) * 40
    $out = Join-Path $LogDir "w$w.out.log"
    $err = Join-Path $LogDir "w$w.err.log"
    Remove-Item $out, $err -Force -ErrorAction SilentlyContinue

    $args = @(
        $Launcher,
        "--worker", "w$w",
        "--offset", "$offset",
        "--window-x", "$x",
        "--window-y", "$y"
    )
    if ($Limit -gt 0) {
        $args += @("--limit", "$Limit")
    }
    if ($Month -ge 1 -and $Month -le 12) {
        $args += @("--month", "$Month")
    }
    if ($File) {
        $args += @("--file", $File)
    }

    Write-Host "Start w$w offset=$offset limit=$Limit"
    $p = Start-Process -FilePath $Py -ArgumentList $args `
        -WorkingDirectory $BrowserAutoRoot `
        -RedirectStandardOutput $out `
        -RedirectStandardError $err `
        -PassThru
    $procs += [pscustomobject]@{ Worker = "w$w"; Offset = $offset; Pid = $p.Id }
    Start-Sleep -Milliseconds 1000
}

$procs | Format-Table -AutoSize
$procs | ConvertTo-Json | Set-Content (Join-Path $LogDir "pids.json") -Encoding utf8
Write-Host "Логи: $LogDir"
Write-Host "CSV результаты появятся в: $DataDir"
