<#
.SYNOPSIS
    Runs the Excel E2E in all four write modes and prints a one-line result per
    mode plus the environment's COM-retry total (the antivirus fingerprint).

.DESCRIPTION
    Used to compare runs under different antivirus states (see
    test\Set-Antivirus.ps1 and AGENTS.md). Each mode writes its full log to
    %TEMP%\e2e-<Tag>-<mode>.log; test\Compare-ExcelE2E.ps1 diffs two tags.

    This is NOT part of CI (it needs Excel). Run it locally, e.g. once with the
    AV active (default tag) and once with it paused.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E-Matrix.ps1 -Tag avon
    # ... pause the AV ...
    powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E-Matrix.ps1 -Tag avoff
    powershell -ExecutionPolicy Bypass -File test\Compare-ExcelE2E.ps1 -TagA avon -TagB avoff
#>
param(
    [string]$Tag = 'run',
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$RedisHost = 'localhost:6379',
    [string]$RedisCli = 'docker exec my-redis redis-cli'
)

$ErrorActionPreference = 'Stop'
$e2e = Join-Path $PSScriptRoot 'Run-ExcelE2E.ps1'
if (-not (Test-Path -LiteralPath $e2e)) { throw "E2E script not found: $e2e" }

$modes = @(
    @{ Name = 'default'; Args = @() },
    @{ Name = 'sync';    Args = @('-SyncWrite', 'sync') },
    @{ Name = 'ffall';   Args = @('-SyncWrite', 'fireforget-all') },
    @{ Name = 'async';   Args = @('-AsyncWrites') }
)

foreach ($mode in $modes) {
    $log = Join-Path $env:TEMP ("e2e-{0}-{1}.log" -f $Tag, $mode.Name)
    Write-Host ("=== MODE {0} [{1}] ===" -f $mode.Name, $Tag) -ForegroundColor Cyan
    & powershell -ExecutionPolicy Bypass -File $e2e `
        -RepoRoot $RepoRoot -RedisHost $RedisHost -RedisCli $RedisCli @($mode.Args) 2>&1 |
        Tee-Object -FilePath $log |
        Select-String 'ALL PASS|FAILURE\(S\)|^FAIL |COM retries:' |
        ForEach-Object { $_.Line.Trim() }
    # Each run leaves its hidden Excel: quit any left behind before the next mode
    # (empty window title = a hidden automation instance, never the user's).
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } |
        ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 3
}
Write-Host ("=== DONE [{0}] ===" -f $Tag) -ForegroundColor Cyan
