<#
  read-config-host.ps1 - prints the server address an existing RedisExcel.json
  points at, in either shape:

    v1.6      Default -> Servers[<name>] (a connection string or { host })
    pre-1.6   RTD.host / UDF.host

  Why: on an upgrade the installer keeps a file the add-in accepts, but a pre-1.6
  file is rejected wholesale (every cell shows #CONFIG), so it is backed up and
  replaced. Reading the old address first keeps that migration from silently
  moving the user to localhost.

  Writes "HOST=<value>" to -Out (nothing when there is no usable address) and
  always exits 0: an unreadable or invalid file simply means "unknown", and the
  installer falls back to its own default.

  Usage:
    powershell -ExecutionPolicy Bypass -File read-config-host.ps1 -Config C:\RedisExcel\RedisExcel.json -Out %TEMP%\host.txt
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$Config,
  [Parameter(Mandatory = $true)][string]$Out
)

$ErrorActionPreference = 'Continue'

# Comments (// and /* */) are allowed in the file, so strip them string-aware
# before parsing - the same tolerance the add-in has.
function Strip-JsonComments([string]$text) {
  $builder = New-Object System.Text.StringBuilder
  $inString = $false
  $inLine = $false
  $inBlock = $false
  for ($i = 0; $i -lt $text.Length; $i++) {
    $c = $text[$i]
    $next = if ($i + 1 -lt $text.Length) { $text[$i + 1] } else { [char]0 }
    if ($inLine) { if ($c -eq "`n") { $inLine = $false; [void]$builder.Append($c) }; continue }
    if ($inBlock) { if ($c -eq '*' -and $next -eq '/') { $inBlock = $false; $i++ }; continue }
    if ($inString) {
      [void]$builder.Append($c)
      if ($c -eq '\') { if ($next -ne [char]0) { [void]$builder.Append($next); $i++ } }
      elseif ($c -eq '"') { $inString = $false }
      continue
    }
    if ($c -eq '/' -and $next -eq '/') { $inLine = $true; $i++; continue }
    if ($c -eq '/' -and $next -eq '*') { $inBlock = $true; $i++; continue }
    if ($c -eq '"') { $inString = $true }
    [void]$builder.Append($c)
  }
  return $builder.ToString()
}

$host_ = ''
try {
  if (Test-Path -LiteralPath $Config) {
    $json = (Strip-JsonComments ([System.IO.File]::ReadAllText($Config))) | ConvertFrom-Json
    $servers = $json.Servers
    if ($servers) {
      $entry = $null
      if ($json.Default) {
        $entry = $servers.PSObject.Properties |
          Where-Object { $_.Name -eq [string]$json.Default } | Select-Object -First 1
      }
      if (-not $entry) { $entry = $servers.PSObject.Properties | Select-Object -First 1 }
      if ($entry) {
        if ($entry.Value -is [string]) { $host_ = $entry.Value }
        elseif ($entry.Value.host) { $host_ = [string]$entry.Value.host }
      }
    }
    if (-not $host_ -and $json.UDF -and $json.UDF.host) { $host_ = [string]$json.UDF.host }
    if (-not $host_ -and $json.RTD -and $json.RTD.host) { $host_ = [string]$json.RTD.host }
  }
}
catch { $host_ = '' }

$line = if ($host_ -ne '') { 'HOST=' + $host_ } else { '' }
Set-Content -LiteralPath $Out -Value $line -Encoding ASCII
if ($host_ -ne '') { Write-Host ("[host] " + $host_) } else { Write-Host '[host] none found' }
exit 0
