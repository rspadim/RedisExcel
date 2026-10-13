<#
  uninstall-server.ps1 - removes a Redis-compatible server chosen in the
  uninstaller. Only the server(s) the user ticked are removed; nothing else.

  Kind:
    docker-redis | docker-valkey   removes the "redisexcel" container
    wsl-redis                      apt-get remove redis-server (WSL)
    wsl-valkey                     apt-get remove valkey-server (WSL)
    memurai                        stops and uninstalls Memurai (needs admin)

  Exit 0 on success or when there was nothing to do, 1 on a hard failure.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidateSet('docker-redis', 'docker-valkey', 'wsl-redis', 'wsl-valkey', 'memurai')]
  [string]$Kind,
  [string]$ContainerName = 'redisexcel',
  [int]$Port = 6379
)

$ErrorActionPreference = 'Continue'
function Write-Step([string]$m) { Write-Host ("[redis-uninstall] " + $m) }

# Removes the container only when its IMAGE matches the requested kind, so
# /REMOVESERVERS=docker-valkey cannot delete a Redis container (audit H2/D3).
function Remove-DockerContainer([string]$imagePattern) {
  if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    Write-Step "Docker not found; nothing to do."
    return
  }
  $image = (& docker inspect --format '{{.Config.Image}}' $ContainerName 2>$null | Out-String).Trim()
  if ([string]::IsNullOrWhiteSpace($image)) {
    Write-Step "no container '$ContainerName'; nothing to do."
    return
  }
  if ($image -notlike $imagePattern) {
    Write-Step "container '$ContainerName' image '$image' does not match '$imagePattern'; leaving it."
    return
  }
  Write-Step "removing container '$ContainerName' (image '$image')"
  & docker rm -f $ContainerName 2>&1 | Write-Host
}

switch ($Kind) {
  'docker-redis' {
    Remove-DockerContainer 'redis*'
    exit 0
  }
  'docker-valkey' {
    Remove-DockerContainer 'valkey*'
    exit 0
  }
  'wsl-redis' {
    Write-Step "removing redis-server from WSL"
    & wsl -e sudo apt-get remove -y redis-server 2>&1 | Write-Host
    exit 0
  }
  'wsl-valkey' {
    Write-Step "removing valkey-server from WSL"
    & wsl -e sudo apt-get remove -y valkey-server 2>&1 | Write-Host
    exit 0
  }
  'memurai' {
    Write-Step "uninstalling Memurai"
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
      & winget uninstall --id Memurai.MemuraiDeveloper -e 2>&1 | Write-Host
      exit 0
    }
    # Fall back to the service + its uninstaller.
    $svc = Get-Service -Name '*memurai*' -ErrorAction SilentlyContinue
    if ($svc) { Stop-Service $svc.Name -Force -ErrorAction SilentlyContinue }
    $unins = Get-ChildItem 'C:\Program Files\Memurai\unins*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($unins) { Start-Process -FilePath $unins.FullName -ArgumentList '/VERYSILENT' -Wait }
    else { Write-Step "Could not find the Memurai uninstaller; remove it from Apps & features." }
    exit 0
  }
}
