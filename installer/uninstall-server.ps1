<#
  uninstall-server.ps1 - removes a Redis-compatible server chosen in the
  uninstaller. Only the server(s) the user ticked are removed; nothing else.

  Kind:
    docker-redis | docker-valkey   removes the "redisexcel" container
    wsl-redis                      apt-get remove redis-server (WSL)
    wsl-valkey                     apt-get remove valkey-server (WSL)
    memurai                        stops and uninstalls Memurai (needs admin)

  Exit 0 on success or when there was nothing to do, 1 when the removal did NOT
  take effect. The result is verified by DETECTION - the same checks
  detect-servers.ps1 uses - never by the removal tool's own exit code: winget
  exits non-zero for "already gone", apt-get exits 0 for "not installed, so not
  removed", and a failing `docker rm -f` only writes to stderr. Without the
  check the script could not fail, the uninstaller's "could not remove" message
  was unreachable, and a server the user believes is gone could stay behind.
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
function Have-Command([string]$name) { return [bool](Get-Command $name -ErrorAction SilentlyContinue) }

# ---- detection (mirrors detect-servers.ps1) --------------------------------
function Test-DockerContainer([string]$name) {
  if (-not (Have-Command docker)) { return $false }
  $image = (& docker inspect --format '{{.Config.Image}}' $name 2>$null | Out-String).Trim()
  # The exit code is the authoritative "no such container" signal (mirrors
  # detect-servers.ps1): `docker inspect` without a successful lookup can still
  # print to stdout (e.g. an empty JSON array), which must not read as "present".
  if ($LASTEXITCODE -ne 0) { return $false }
  return -not [string]::IsNullOrWhiteSpace($image)
}

function Test-WslServer([string]$command) {
  if (-not (Have-Command wsl)) { return $false }
  try {
    $out = (& wsl -e sh -c ("command -v " + $command + " >/dev/null 2>&1 && echo YES || echo NO") 2>$null | Out-String)
    return ($out -match 'YES')
  }
  catch { return $false }
}

function Test-Memurai {
  return [bool](Get-Service -Name '*memurai*' -ErrorAction SilentlyContinue)
}

# ---- removal (each returns $true when the target is gone afterwards) --------
# Removes the container only when its IMAGE matches the requested kind, so
# /REMOVESERVERS=docker-valkey cannot delete a Redis container (audit H2/D3).
function Remove-DockerContainer([string]$imagePattern) {
  if (-not (Have-Command docker)) {
    Write-Step "Docker not found; nothing to do."
    return $true
  }
  $image = (& docker inspect --format '{{.Config.Image}}' $ContainerName 2>$null | Out-String).Trim()
  if ([string]::IsNullOrWhiteSpace($image)) {
    Write-Step "no container '$ContainerName'; nothing to do."
    return $true
  }
  if ($image -notlike $imagePattern) {
    Write-Step "container '$ContainerName' image '$image' does not match '$imagePattern'; leaving it."
    return $true
  }
  Write-Step "removing container '$ContainerName' (image '$image')"
  & docker rm -f $ContainerName 2>&1 | Write-Host
  if (Test-DockerContainer $ContainerName) {
    Write-Step "container '$ContainerName' is STILL present after 'docker rm -f'."
    return $false
  }
  Write-Step "container '$ContainerName' removed."
  return $true
}

function Remove-WslServer([string]$command, [string]$package) {
  if (-not (Have-Command wsl)) {
    Write-Step "WSL not found; nothing to do."
    return $true
  }
  if (-not (Test-WslServer $command)) {
    Write-Step "$command is not installed in WSL; nothing to do."
    return $true
  }
  Write-Step "removing $package from WSL"
  & wsl -e sudo apt-get remove -y $package 2>&1 | Write-Host
  if (Test-WslServer $command) {
    Write-Step "$command is STILL installed in WSL after 'apt-get remove'."
    return $false
  }
  Write-Step "$command removed from WSL."
  return $true
}

function Remove-Memurai {
  if (-not (Test-Memurai)) {
    Write-Step "no Memurai service found; nothing to do."
    return $true
  }
  Write-Step "uninstalling Memurai"
  if (Have-Command winget) {
    & winget uninstall --id Memurai.MemuraiDeveloper -e 2>&1 | Write-Host
  }
  else {
    # Fall back to the service + its uninstaller.
    $svc = Get-Service -Name '*memurai*' -ErrorAction SilentlyContinue
    if ($svc) { Stop-Service $svc.Name -Force -ErrorAction SilentlyContinue }
    $unins = Get-ChildItem 'C:\Program Files\Memurai\unins*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($unins) { Start-Process -FilePath $unins.FullName -ArgumentList '/VERYSILENT' -Wait }
    else { Write-Step "Could not find the Memurai uninstaller." }
  }
  if (Test-Memurai) {
    Write-Step "the Memurai service is STILL present after the uninstall attempt (needs admin?)."
    return $false
  }
  Write-Step "Memurai removed."
  return $true
}

$ok = $true
switch ($Kind) {
  'docker-redis'  { $ok = Remove-DockerContainer 'redis*' }
  'docker-valkey' { $ok = Remove-DockerContainer 'valkey*' }
  'wsl-redis'     { $ok = Remove-WslServer 'redis-server' 'redis-server' }
  'wsl-valkey'    { $ok = Remove-WslServer 'valkey-server' 'valkey-server' }
  'memurai'       { $ok = Remove-Memurai }
}

if ($ok) { exit 0 }
Write-Step "the removal did not take effect; remove $Kind manually if needed."
exit 1
