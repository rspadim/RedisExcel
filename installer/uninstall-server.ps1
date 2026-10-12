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

switch ($Kind) {
  'docker-redis' {
    $docker = Get-Command docker -ErrorAction SilentlyContinue
    if (-not $docker) { Write-Step "Docker not found; nothing to do."; exit 0 }
    Write-Step "removing container '$ContainerName'"
    & docker rm -f $ContainerName 2>&1 | Write-Host
    exit 0
  }
  'docker-valkey' {
    $docker = Get-Command docker -ErrorAction SilentlyContinue
    if (-not $docker) { Write-Step "Docker not found; nothing to do."; exit 0 }
    Write-Step "removing container '$ContainerName'"
    & docker rm -f $ContainerName 2>&1 | Write-Host
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
