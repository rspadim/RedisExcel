<#
  test-connection.ps1 - quick Redis PING used by the installer wizard.
  Prints the result and exits 0 when the server answers PONG, 1 otherwise.
  No windows/dialogs: the wizard shows the message.

  Example:
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:6379
#>
[CmdletBinding()]
param(
  [string]$ServerHost = 'localhost',
  [int]$Port = 6379,
  [string]$User = '',
  [string]$Password = ''
)

$ErrorActionPreference = 'Stop'

$hostName = $ServerHost
$port = $Port
if ($ServerHost -match '^\s*([^,:;]+)(?::(\d+))?') {
  $hostName = $Matches[1]
  if ($Matches[2]) { $port = [int]$Matches[2] }
}

try {
  $client = New-Object System.Net.Sockets.TcpClient
  $client.Connect($hostName, $port)
  $stream = $client.GetStream()
  $writer = New-Object System.IO.StreamWriter($stream)
  $writer.NewLine = "`r`n"
  $writer.AutoFlush = $true
  $reader = New-Object System.IO.StreamReader($stream)

  if ($Password -ne '') {
    if ($User -ne '') { $writer.WriteLine("AUTH $User $Password") } else { $writer.WriteLine("AUTH $Password") }
    $auth = $reader.ReadLine()
    if ($auth -notmatch '^\+OK') {
      Write-Host "AUTH failed: $auth"
      $client.Close()
      exit 1
    }
  }

  $writer.WriteLine('PING')
  $pong = $reader.ReadLine()
  $client.Close()
  if ($pong -match 'PONG') { Write-Host "PONG"; exit 0 }
  Write-Host "unexpected reply: $pong"
  exit 1
} catch {
  Write-Host ("no answer: " + $_.Exception.Message)
  exit 1
}
