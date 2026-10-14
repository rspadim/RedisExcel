#Requires -Version 5.1
<#
    Run-Tests.ps1 - the single-command local test run for RedisExcel.

    It builds once and then runs the independent test layers off that one build
    (--no-build), in PARALLEL by default, so no layer pays for a second
    restore/compile and a slow layer does not block the others.

    Layers:
      unit      - no server. Always run.
      smoke     - needs a Redis endpoint. Its own container unless -RedisHost is
                  given.
      liveness  - needs a Redis endpoint; --quick --skip-restart by default (the
                  CI set). -LivenessFull drops both flags for the destructive run.
      e2e       - NOT started unless -WithE2E (needs Excel + Redis).

    The layers keep their own semantics: this is a convenience wrapper, it never
    skips or weakens an assertion.

    Usage:
      powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1
      ... -NoBuild                 # reuse the current build (fastest re-run)
      ... -WithE2E                 # add the Excel end-to-end layer
      ... -SkipSmoke -SkipLiveness # unit only
      ... -LivenessFull            # the full destructive liveness run
      ... -RedisHost host:port     # use an existing server (no container)
      ... -ContainerName rs-rt     # name for the disposable redis:7-alpine
      ... -Sequential              # run one layer at a time (easier to read)
#>
[CmdletBinding()]
param(
  [switch]$NoBuild,
  [switch]$WithE2E,
  [switch]$SkipSmoke,
  [switch]$SkipLiveness,
  [switch]$LivenessFull,
  [switch]$Sequential,
  [string]$RedisHost,
  [string]$ContainerName = 'rs-runtests',
  [int]$Port = 6399
)

$ErrorActionPreference = 'Continue'   # native tools (docker/dotnet) write to stderr; every
                                      # critical step below checks $LASTEXITCODE explicitly.
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not (Test-Path (Join-Path $repoRoot 'RedisExcel.sln'))) {
  throw "Run-Tests.ps1: '$repoRoot' does not look like the repo root (no RedisExcel.sln)"
}
Push-Location $repoRoot
$script:Failures = @()
$script:StartedContainer = $false

function Write-Layer([string]$name, [string]$message) {
  Write-Host ("[{0,-8}] {1}" -f $name, $message)
}

function Invoke-Layer([string]$name, [scriptblock]$body) {
  Write-Layer $name 'start'
  $started = Get-Date
  try {
    & $body
    $seconds = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
    if ($LASTEXITCODE -ne 0 -and $null -ne $LASTEXITCODE) {
      $script:Failures += $name
      Write-Layer $name "FAIL (exit $LASTEXITCODE, ${seconds}s)"
    }
    else {
      Write-Layer $name "ok (${seconds}s)"
    }
  }
  catch {
    $script:Failures += $name
    Write-Layer $name ("FAIL: " + $_.Exception.Message)
  }
}

function Ensure-Redis {
  if ($RedisHost) { return $RedisHost }
  $docker = Get-Command docker -ErrorAction SilentlyContinue
  if (-not $docker) {
    throw 'no -RedisHost given and Docker is not available: start a Redis and pass -RedisHost host:port'
  }
  & docker rm -f $ContainerName *> $null
  & docker run -d --name $ContainerName -p "${Port}:6379" redis:7-alpine *> $null
  if ($LASTEXITCODE -ne 0) { throw "could not start the '$ContainerName' container" }
  $script:StartedContainer = $true
  foreach ($attempt in 1..20) {
    & docker exec $ContainerName redis-cli ping *> $null
    if ($LASTEXITCODE -eq 0) { return "127.0.0.1:$Port" }
    Start-Sleep -Milliseconds 500
  }
  throw "the '$ContainerName' container never answered PING"
}

try {
  if (-not $NoBuild) {
    Write-Layer build 'dotnet build RedisExcel.sln -c Release'
    & dotnet build RedisExcel.sln -c Release --nologo
    if ($LASTEXITCODE -ne 0) { throw 'the build failed; fix it before running the layers' }
  }

  $needsRedis = (-not $SkipSmoke) -or (-not $SkipLiveness) -or $WithE2E
  $redis = $null
  if ($needsRedis) { $redis = Ensure-Redis }

  $layers = New-Object System.Collections.Generic.List[object]
  $layers.Add(@{ Name = 'unit'; Body = {
      & dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release --nologo --no-build
    } })

  if (-not $SkipSmoke) {
    $smoke = @{ Name = 'smoke'; Body = {
        & dotnet run --project test\SmokeTests -c Release --no-build -- $redis
      } }
    $layers.Add($smoke)
  }

  if (-not $SkipLiveness) {
    $flags = if ($LivenessFull) { @() } else { @('--quick', '--skip-restart') }
    $liveness = @{ Name = 'liveness'; Body = {
        $argsList = @('run', '--project', 'test\LivenessTests', '-c', 'Release', '--no-build', '--', "$redis,abortConnect=False") + $flags
        & dotnet $argsList
      } }
    $layers.Add($liveness)
  }

  if ($WithE2E) {
    $e2e = @{ Name = 'e2e'; Body = {
        & powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 -RedisHost $redis
      } }
    $layers.Add($e2e)
  }

  if ($Sequential) {
    foreach ($layer in $layers) { Invoke-Layer $layer.Name $layer.Body }
  }
  else {
    $jobs = foreach ($layer in $layers) {
      Start-Job -Name $layer.Name -ScriptBlock $layer.Body
    }
    foreach ($job in $jobs) {
      Wait-Job $job | Out-Null
      $output = Receive-Job $job
      $output | ForEach-Object { Write-Host $_ }
      if ($job.State -ne 'Completed') {
        $script:Failures += $job.Name
        Write-Layer $job.Name ("FAIL: job state " + $job.State)
      }
      elseif ($job.ChildJobs[0].JobStateInfo.Reason) {
        $script:Failures += $job.Name
        Write-Layer $job.Name ("FAIL: " + $job.ChildJobs[0].JobStateInfo.Reason.Message)
      }
      else {
        Write-Layer $job.Name 'ok'
      }
      Remove-Job $job
    }
  }
}
finally {
  if ($script:StartedContainer) {
    Write-Layer cleanup "removing the '$ContainerName' container"
    & docker rm -f $ContainerName *> $null
  }
  Pop-Location
}

Write-Host ''
if ($script:Failures.Count -gt 0) {
  Write-Host ("RESULT: FAIL - " + ($script:Failures -join ', ')) -ForegroundColor Red
  exit 1
}
Write-Host 'RESULT: ALL PASS' -ForegroundColor Green
exit 0
