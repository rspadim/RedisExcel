<#
  detect-servers.ps1 - lists Redis-compatible servers found on this machine, so
  the RedisExcel uninstaller can offer to remove the ones it created.

  Prints one line per candidate:  kind|label|present
  (present = 1 when found, 0 otherwise). The uninstaller parses this.

  Kinds:
    docker-redis    a container named "redisexcel" running redis
    docker-valkey   a container named "redisexcel" running valkey
    wsl-redis       redis-server installed in a WSL distro
    wsl-valkey      valkey-server installed in a WSL distro
    memurai         the Memurai Windows service
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

function Have-Command([string]$name) {
  return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

# ---- Docker containers named "redisexcel" ---------------------------------
$dockerOk = $false
if (Have-Command docker) {
  try { & docker version --format '{{.Server.Version}}' *> $null; $dockerOk = ($LASTEXITCODE -eq 0) } catch { $dockerOk = $false }
}

function Docker-Image([string]$name) {
  if (-not $dockerOk) { return '' }
  try {
    $img = (& docker inspect --format '{{.Config.Image}}' $name 2>$null)
    if ($LASTEXITCODE -eq 0) { return ($img | Out-String).Trim() } else { return '' }
  } catch { return '' }
}

$cRedisImg = Docker-Image 'redisexcel'
$hasDockerRedis = ($cRedisImg -ne '' -and $cRedisImg -like 'redis*')
$hasDockerValkey = ($cRedisImg -ne '' -and $cRedisImg -like 'valkey*')
# If the image string is unknown but the container exists, still offer redisel.
if ($cRedisImg -ne '' -and -not $hasDockerRedis -and -not $hasDockerValkey) { $hasDockerRedis = $true }

Write-Output ("docker-redis|Redis (Docker container 'redisexcel')|" + $(if ($hasDockerRedis) { '1' } else { '0' }))
Write-Output ("docker-valkey|Valkey (Docker container 'redisexcel')|" + $(if ($hasDockerValkey) { '1' } else { '0' }))

# ---- WSL servers ----------------------------------------------------------
function Wsl-Has([string]$cmd) {
  if (-not (Have-Command wsl)) { return $false }
  try {
    $out = (& wsl -e sh -c ("command -v " + $cmd + " >/dev/null 2>&1 && echo YES || echo NO") 2>$null | Out-String)
    return ($out -match 'YES')
  } catch { return $false }
}

$hasWslRedis = Wsl-Has 'redis-server'
$hasWslValkey = Wsl-Has 'valkey-server'
Write-Output ("wsl-redis|Redis (WSL2)|" + $(if ($hasWslRedis) { '1' } else { '0' }))
Write-Output ("wsl-valkey|Valkey (WSL2)|" + $(if ($hasWslValkey) { '1' } else { '0' }))

# ---- Memurai service ------------------------------------------------------
$memurai = Get-Service -Name '*memurai*' -ErrorAction SilentlyContinue
Write-Output ("memurai|Memurai (Windows service)|" + $(if ($memurai) { '1' } else { '0' }))
