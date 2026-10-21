<#
  test-connection.ps1 - quick Redis PING used by the installer wizard.
  Prints the result and exits 0 when the server answers PONG, 1 otherwise.
  No windows/dialogs: the wizard shows the message.

  AUTH is sent as a RESP command (bulk strings), so a password with spaces or
  other special characters works; the inline "AUTH <pass>" form Redis also
  accepts cannot carry a space. -PasswordFile takes precedence over -Password:
  the wizard writes the secret to a private temp file so a password containing a
  quote or a space never has to survive argument quoting (and never shows up in a
  process command line).

  The socket connect is BOUNDED: an unreachable host must not freeze the wizard
  on the OS SYN timeout.

  Exit codes are STABLE reason codes so a caller can show its own (localized)
  message instead of this script's text: 0 = PONG, 2 = unknown host (DNS),
  3 = timeout/unreachable, 4 = connection refused or other socket error,
  5 = AUTH failed, 6 = unexpected reply. The human detail is written after the
  first line; it is free-form (may be non-ASCII) and meant for the log, not for
  a UI label - a console redirect re-encodes it and would mangle accents.

  Examples:
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:6379
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:6379 -PasswordFile C:\Temp\pw.txt
#>
[CmdletBinding()]
param(
  [string]$ServerHost = 'localhost',
  [int]$Port = 6379,
  [string]$User = '',
  [string]$Password = '',
  [string]$PasswordFile = ''
)

$ErrorActionPreference = 'Stop'
$ConnectTimeoutMs = 5000

# The file content is used verbatim (no Trim): a password may contain spaces.
if ($PasswordFile -ne '' -and (Test-Path -LiteralPath $PasswordFile)) {
  $Password = [System.IO.File]::ReadAllText($PasswordFile)
}

$hostName = $ServerHost
$port = $Port
if ($ServerHost -match '^\s*([^,:;]+)(?::(\d+))?') {
  $hostName = $Matches[1]
  if ($Matches[2]) { $port = [int]$Matches[2] }
}

function Connect-Redis([string]$h, [int]$p) {
  $client = New-Object System.Net.Sockets.TcpClient
  $connect = $client.ConnectAsync($h, $p)
  if (-not $connect.Wait($ConnectTimeoutMs)) {
    $client.Close()
    # 3 = no answer in time (unreachable host, firewall, wrong port)
    exit 3
  }
  return $client
}

# A failed Task.Wait surfaces as an AggregateException, whose own message says
# nothing ("Exception calling Wait with 1 argument(s)"). The innermost exception
# carries the reason the user needs ("no connection could be made because the
# target machine actively refused it").
function Get-ErrorMessage($record) {
  $ex = $record.Exception
  while ($ex -and $ex.InnerException) { $ex = $ex.InnerException }
  if ($ex) { return $ex.Message }
  return $record.Exception.Message
}

# RESP array of bulk strings: the only form that can carry spaces and binary
# characters in a Redis command.
function Send-Command($stream, [string[]]$parts) {
  $builder = New-Object System.Text.StringBuilder
  [void]$builder.Append('*' + $parts.Count + "`r`n")
  foreach ($part in $parts) {
    $byteCount = [System.Text.Encoding]::UTF8.GetByteCount($part)
    [void]$builder.Append('$' + $byteCount + "`r`n")
    [void]$builder.Append($part + "`r`n")
  }
  $payload = [System.Text.Encoding]::UTF8.GetBytes($builder.ToString())
  $stream.Write($payload, 0, $payload.Length)
  $stream.Flush()
}

try {
  $client = Connect-Redis $hostName $port
  $stream = $client.GetStream()
  $reader = New-Object System.IO.StreamReader($stream)

  if ($Password -ne '') {
    if ($User -ne '') { Send-Command $stream @('AUTH', $User, $Password) }
    else { Send-Command $stream @('AUTH', $Password) }
    $auth = $reader.ReadLine()
    if ($auth -notmatch '^\+OK') {
      # 5 = the server answered but rejected the credentials
      Write-Host "AUTH failed: $auth"
      $client.Close()
      exit 5
    }
  }

  Send-Command $stream @('PING')
  $pong = $reader.ReadLine()
  $client.Close()
  if ($pong -match 'PONG') { Write-Host "PONG"; exit 0 }
  # 6 = something answered, but not with a PONG
  Write-Host "unexpected reply: $pong"
  exit 6
} catch {
  $message = Get-ErrorMessage $_
  Write-Host ("no answer: " + $message)
  # 2 = the NAME does not resolve (a typo in the host); 4 = the name resolves
  # but nothing accepts the connection (server down, wrong port, firewall).
  # The classification uses the SocketError CODE, never the message text: the
  # text is localized by Windows and a pattern match on it breaks per locale.
  $hostNotFound = $false
  $e = $_.Exception
  while ($e -and -not $hostNotFound) {
    $se = $e -as [System.Net.Sockets.SocketException]
    if ($se -and ($se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound -or
                  $se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::NoData)) {
      $hostNotFound = $true
    }
    $e = $e.InnerException
  }
  if ($hostNotFound) { exit 2 }
  exit 4
}
