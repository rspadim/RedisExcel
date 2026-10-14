#Requires -Version 5.1
<#
    Run-Tests.ps1 - the single-command local test run for RedisExcel.

    It builds once and then runs the independent test layers off that one build
    (--no-build), in PARALLEL by default, so no layer pays for a second
    restore/compile and a slow layer does not block the others.

    Layers:
      unit      - no server. Always run.
      smoke     - needs a Redis endpoint.
      liveness  - needs a Redis endpoint; --quick --skip-restart by default (the
                  CI set). -LivenessFull drops both flags for the destructive run.
      e2e       - NOT started unless -WithE2E (needs Excel + Redis).

    Without -RedisHost the script starts its own disposable redis:7-alpine
    container and removes it afterwards. The layers keep their own semantics:
    this is a convenience wrapper, it never skips or weakens an assertion.

    Usage:
      powershell -ExecutionPolicy Bypass -File test\TestSupport\Run-Tests.ps1
      ... -NoBuild                 # reuse the current build (fastest re-run)
      ... -WithE2E                 # add the Excel end-to-end layer
      ... -SkipSmoke -SkipLiveness # unit only
      ... -LivenessFull            # the full destructive liveness run
      ... -RedisHost host:port     # use an existing server (no container)
      ... -ContainerName rs-rt     # name for the disposable redis:7-alpine
      ... -Port 6399               # host port for the disposable container
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

# Native tools (docker/dotnet) write to stderr; every critical step below checks
# $LASTEXITCODE explicitly instead of relying on the error preference.
$ErrorActionPreference = 'Continue'
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

function Format-Seconds([datetime]$started) {
  return [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
}

function Ensure-Redis {
  if ($RedisHost) { return $RedisHost }
  if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    throw 'no -RedisHost given and Docker is not available: start a Redis and pass -RedisHost host:port'
  }
  & docker rm -f $ContainerName *> $null
  & docker run -d --name $ContainerName -p "${Port}:6379" redis:7-alpine *> $null
  if ($LASTEXITCODE -ne 0) { throw "could not start the '$ContainerName' container on port $Port" }
  $script:StartedContainer = $true
  foreach ($attempt in 1..20) {
    & docker exec $ContainerName redis-cli ping *> $null
    if ($LASTEXITCODE -eq 0) { return "127.0.0.1:$Port" }
    Start-Sleep -Milliseconds 500
  }
  throw "the '$ContainerName' container never answered PING"
}

# One layer as a native command line, run in $repoRoot. The exit code is
# returned EXPLICITLY: inside a job the command's failure does not change the
# job state, so a state check alone would report a false pass.
function Invoke-LayerCommand([string]$command) {
  $output = & ([scriptblock]::Create($command)) 2>&1
  return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
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
  $layers.Add(@{ Name = 'unit'; Command =
      'dotnet test test\RedisExcel.Tests\RedisExcel.Tests.csproj -c Release --nologo --no-build' })

  if (-not $SkipSmoke) {
    $layers.Add(@{ Name = 'smoke'; Command =
        "dotnet run --project test\SmokeTests -c Release --no-build -- `"$redis,abortConnect=False`"" })
  }

  if (-not $SkipLiveness) {
    $flags = if ($LivenessFull) { '' } else { '--quick --skip-restart' }
    $layers.Add(@{ Name = 'liveness'; Command =
        "dotnet run --project test\LivenessTests -c Release --no-build -- `"$redis,abortConnect=False`" $flags" })
  }

  if ($WithE2E) {
    $layers.Add(@{ Name = 'e2e'; Command =
        "powershell -NoProfile -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 -RedisHost `"$redis`"" })
  }

  $started = @{}
  foreach ($layer in $layers) { $started[$layer.Name] = Get-Date }

  if ($Sequential) {
    foreach ($layer in $layers) {
      Write-Layer $layer.Name 'start'
      $result = Invoke-LayerCommand $layer.Command
      $result.Output | ForEach-Object { Write-Host $_ }
      if ($result.ExitCode -ne 0) {
        $script:Failures += $layer.Name
        Write-Layer $layer.Name ("FAIL (exit {0}, {1}s)" -f $result.ExitCode, (Format-Seconds $started[$layer.Name]))
      }
      else {
        Write-Layer $layer.Name ("ok ({0}s)" -f (Format-Seconds $started[$layer.Name]))
      }
    }
  }
  else {
    $jobs = foreach ($layer in $layers) {
      Start-Job -Name $layer.Name -ArgumentList $repoRoot, $layer.Command -ScriptBlock {
        param($root, $command)
        Set-Location $root
        $output = & ([scriptblock]::Create($command)) 2>&1
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($output | Out-String) }
      }
    }
    foreach ($job in $jobs) {
      Wait-Job $job | Out-Null
      # The job body writes exactly one result object; anything else means the
      # body itself blew up before the command ran.
      $payload = Receive-Job $job | Where-Object { $_ -and $_.PSObject.Properties['ExitCode'] } | Select-Object -Last 1
      $seconds = Format-Seconds $started[$job.Name]
      if ($null -eq $payload) {
        $script:Failures += $job.Name
        Write-Layer $job.Name ("FAIL: the job produced no result (state {0})" -f $job.State)
      }
      else {
        if ($payload.Output) { Write-Host $payload.Output.TrimEnd() }
        if ($payload.ExitCode -ne 0) {
          $script:Failures += $job.Name
          Write-Layer $job.Name ("FAIL (exit {0}, {1}s)" -f $payload.ExitCode, $seconds)
        }
        else {
          Write-Layer $job.Name ("ok ({0}s)" -f $seconds)
        }
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
