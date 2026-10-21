<#
  install-server.ps1 - optional Redis/Valkey server setup for the RedisExcel
  installer. Called by the wizard (or by hand) to bring a server up on the
  chosen port and leave it running.

  Kinds:
    docker-redis    Docker Desktop container "redisexcel" running redis:latest
    docker-valkey   Docker Desktop container "redisexcel" running valkey/valkey:latest
    wsl-redis       Redis inside the default WSL2 distro
    wsl-valkey      Valkey inside the default WSL2 distro
    memurai         Memurai (native Redis-compatible Windows service; needs admin)
    none            do nothing

  Use -Image to pin a specific tag (e.g. -Image redis:7-alpine) instead of
  latest.

  Exits 0 on success (or when the server already answers), 3 when the chosen
  method is not available on this machine, 1 on a hard failure. The installer
  reports the exit code but never fails the add-in install because of it.

  Examples:
    powershell -ExecutionPolicy Bypass -File install-server.ps1 -Kind docker-valkey -Port 6379
    powershell -ExecutionPolicy Bypass -File install-server.ps1 -Kind wsl-redis
    powershell -ExecutionPolicy Bypass -File install-server.ps1 -Kind memurai
#>
[CmdletBinding()]
param(
  [ValidateSet('docker-redis', 'docker-valkey', 'wsl-redis', 'wsl-valkey', 'memurai', 'none')]
  [string]$Kind = 'none',
  [int]$Port = 6379,
  [string]$ContainerName = 'redisexcel',
  [string]$Distro = '',
  # Override the container image (default: latest). e.g. -Image redis:7-alpine
  [string]$Image = ''
)

$ErrorActionPreference = 'Stop'
$ExitUnsupported = 3

# Native commands write to stderr for normal conditions (e.g. `docker rm -f` on a
# missing container); do not let that abort the script.
function Invoke-Native([scriptblock]$block) {
  $prev = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try { & $block 2>&1 | ForEach-Object { $_ } } finally { $ErrorActionPreference = $prev }
}

function Write-Step([string]$m) { Write-Host ("[redis-server] " + $m) }

# The console header: this script runs in a visible window during the install,
# and an unidentified window that prints technical lines is the main "is this a
# virus?" trigger. Dot-sourced so a missing banner.ps1 can never break the setup.
$script:BannerLoaded = $false
try {
  $bannerPath = Join-Path $PSScriptRoot 'banner.ps1'
  if (Test-Path -LiteralPath $bannerPath) {
    . $bannerPath
    $script:BannerLoaded = $true
  }
} catch { $script:BannerLoaded = $false }

function Show-Banner([string]$action) {
  if ($script:BannerLoaded) {
    try { Write-RedisExcelBanner -Action $action } catch { }
  }
}

function Test-Port([int]$p) {
  try {
    $c = New-Object System.Net.Sockets.TcpClient
    $c.Connect('127.0.0.1', $p)
    $c.Close()
    return $true
  } catch { return $false }
}

function Wait-Port([int]$p, [int]$seconds = 40) {
  $deadline = (Get-Date).AddSeconds($seconds)
  while ((Get-Date) -lt $deadline) {
    if (Test-Port $p) { return $true }
    Start-Sleep -Milliseconds 700
  }
  return (Test-Port $p)
}

function Get-Docker() {
  $d = Get-Command docker -ErrorAction SilentlyContinue
  if (-not $d) { return $null }
  # The CLI exists even when the engine is down; require a live server.
  try {
    $prev = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    & $d.Source version --format '{{.Server.Version}}' *> $null
    $ErrorActionPreference = $prev
    if ($LASTEXITCODE -eq 0) { return $d.Source }
  } catch { } finally { $ErrorActionPreference = 'Stop' }
  return $null
}

function Get-Wsl() {
  $w = Get-Command wsl -ErrorAction SilentlyContinue
  if (-not $w) { return $null }
  return $w.Source
}

function Docker-Up([string]$image, [string]$cliPingContainer) {
  $docker = Get-Docker
  if (-not $docker) {
    Write-Step "Docker is not installed or its engine is not running."
    exit $ExitUnsupported
  }
  if (Test-Port $Port) {
    Write-Step "port $Port is already answering; leaving the existing server in place."
    exit 0
  }
  Write-Step "removing any previous '$ContainerName' container (ignore errors)"
  Invoke-Native { & $docker rm -f $ContainerName } | Out-Null
  Write-Step "starting $image on port $Port (auto-restarts on reboot)"
  $runOut = Invoke-Native { & $docker run -d --name $ContainerName -p "${Port}:6379" --restart unless-stopped $image }
  if ($LASTEXITCODE -ne 0) { Write-Step ("docker run failed: " + (($runOut | Out-String).Trim())); exit 1 }
  if (-not (Wait-Port $Port 60)) { Write-Step "the container started but port $Port never answered"; exit 1 }
  # A real PING through the container's own CLI confirms the protocol, not just the socket.
  try {
    $pong = Invoke-Native { & $docker exec $ContainerName $cliPingContainer ping }
    Write-Step ("PING -> " + (($pong | Out-String).Trim()))
  } catch { }
  Write-Step "done: $image is serving on localhost:$Port"
  exit 0
}

function Wsl-Up([string]$pkg) {
  $wsl = Get-Wsl
  if (-not $wsl) { Write-Step "WSL is not installed."; exit $ExitUnsupported }
  if (Test-Port $Port) { Write-Step "port $Port is already answering; nothing to do."; exit 0 }

  $distroArg = @()
  if ($Distro) { $distroArg = @('-d', $Distro) }
  Write-Step "installing $pkg inside WSL (may prompt for your Linux password)"
  # -e runs a command without a login shell; sudo is needed to install packages.
  & $wsl @distroArg -e sudo apt-get update
  & $wsl @distroArg -e sudo apt-get install -y $pkg
  if ($LASTEXITCODE -ne 0) { Write-Step "apt-get install $pkg failed (exit $LASTEXITCODE)"; exit 1 }
  $binary = if ($pkg -like 'valkey*') { 'valkey-server' } else { 'redis-server' }
  Write-Step "starting $binary in the background on port $Port"
  # Isolate this instance: a dedicated dir avoids the "Can't handle RDB format
  # version" clash when Redis and Valkey share the default /var/lib/redis dump,
  # and the appendonly/save flags keep it out of the way of a service instance.
  $dir = "/var/lib/redisexcel-$Port"
  $start = $binary + " --port $Port --daemonize yes --dir $dir --save '' --appendonly no"
  & $wsl @distroArg -e sudo sh -c "mkdir -p $dir"
  & $wsl @distroArg -e sh -c $start
  if (-not (Wait-Port $Port 30)) {
    Write-Step "the first start failed; retrying with a clean data dir"
    & $wsl @distroArg -e sudo sh -c "rm -f $dir/dump.rdb"
    & $wsl @distroArg -e sh -c $start
    if (-not (Wait-Port $Port 20)) { Write-Step "started $binary but port $Port never answered"; exit 1 }
  }
  Write-Step "done: $pkg is serving on localhost:$Port (inside WSL)"
  exit 0
}

function Memurai-Up() {
  # Prefer winget (signed, versioned); fall back to the Memurai download page.
  $winget = Get-Command winget -ErrorAction SilentlyContinue
  if ($winget) {
    Write-Step "installing Memurai Developer via winget"
    & $winget.Source install --id Memurai.MemuraiDeveloper --accept-source-agreements --accept-package-agreements -e
    if ($LASTEXITCODE -eq 0 -and (Wait-Port $Port 60)) {
      Write-Step "done: Memurai is serving on localhost:$Port"
      exit 0
    }
    Write-Step "winget install did not bring Memurai up (exit $LASTEXITCODE)"
  }
  Write-Step "Could not install Memurai automatically. Download it from https://www.memurai.com and run the installer."
  exit $ExitUnsupported
}

switch ($Kind) {
  'none'          { Show-Banner 'No server requested - nothing to do.'; Write-Step "no server requested."; exit 0 }
  'docker-redis'  { Show-Banner 'Setting up a Redis server in Docker'; Docker-Up $(if ($Image) { $Image } else { 'redis:latest' }) 'redis-cli' }
  'docker-valkey' { Show-Banner 'Setting up a Valkey server in Docker'; Docker-Up $(if ($Image) { $Image } else { 'valkey/valkey:latest' }) 'valkey-cli' }
  'wsl-redis'     { Show-Banner 'Setting up a Redis server in WSL2'; Wsl-Up 'redis-server' }
  'wsl-valkey'    { Show-Banner 'Setting up a Valkey server in WSL2'; Wsl-Up 'valkey-server' }
  'memurai'       { Show-Banner 'Setting up Memurai (native Windows service)'; Memurai-Up }
}
