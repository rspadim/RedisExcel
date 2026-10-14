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
    throw ("no TCP answer within " + [int]($ConnectTimeoutMs / 1000) + "s (unreachable host or firewall?)")
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
      Write-Host "AUTH failed: $auth"
      $client.Close()
      exit 1
    }
  }

  Send-Command $stream @('PING')
  $pong = $reader.ReadLine()
  $client.Close()
  if ($pong -match 'PONG') { Write-Host "PONG"; exit 0 }
  Write-Host "unexpected reply: $pong"
  exit 1
} catch {
  Write-Host ("no answer: " + (Get-ErrorMessage $_))
  exit 1
}
