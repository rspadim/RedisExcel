<#
  verify-connection.ps1 - end-of-install check for RedisExcel.

  * Tests that a Redis-compatible server answers on the given host/port (with
    AUTH when credentials are given, as a RESP command so a password with spaces
    works; the connect is bounded so an unreachable host cannot freeze the
    wizard).
  * With -OpenExcel, opens a visible Excel workbook that loads the installed XLL
    and shows the connection status with a friendly message.

  -PasswordFile takes precedence over -Password: the wizard writes the secret to
  a private temp file, so a password containing a quote or a space never has to
  survive argument quoting (and never shows up in a process command line).

  Exit codes: 0 = Redis answered (and the sheet opened if asked), 1 = no answer.

  Examples:
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\RedisExcel" -ServerHost "localhost" -Port 6379
    powershell -ExecutionPolicy Bypass -File verify-connection.ps1 -AddInDir "C:\RedisExcel" -ServerHost "localhost:6379" -OpenExcel
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$AddInDir,
  [string]$ServerHost = 'localhost',
  [int]$Port = 6379,
  [string]$User = '',
  [string]$Password = '',
  [string]$PasswordFile = '',
  [switch]$OpenExcel,
  [switch]$Quiet,
  # Seconds to keep the console open when the caller runs this in a VISIBLE
  # window. The installer starts it with SW_SHOW, so without a hold the banner
  # and the result would flash and vanish. Excel itself stays open; this only
  # keeps the CONSOLE readable before the script returns.
  [int]$HoldSeconds = 0
)

$ErrorActionPreference = 'Stop'
$ConnectTimeoutMs = 5000
Add-Type -AssemblyName System.Windows.Forms | Out-Null

# Console header (see banner.ps1): this script runs in a visible window during
# the install, and an unidentified window is what makes a user suspect a virus.
# Failure to load it must never break the check itself.
try {
  $bannerPath = Join-Path $PSScriptRoot 'banner.ps1'
  if (Test-Path -LiteralPath $bannerPath) {
    . $bannerPath
    Write-RedisExcelBanner -Action 'Installing and checking RedisExcel' -InstallDir $AddInDir
  }
} catch { }

if ($PasswordFile -ne '' -and (Test-Path -LiteralPath $PasswordFile)) {
  $Password = [System.IO.File]::ReadAllText($PasswordFile)
}

# Keeps the console readable when started from the installer (SW_SHOW). Called
# right before every exit path, so the banner and the result are on screen for
# at least HoldSeconds. With 0 it is a no-op (the redirected wizard call).
function Hold {
  if ($HoldSeconds -gt 0) { Start-Sleep -Seconds $HoldSeconds }
}

function Show([string]$text, [string]$title = 'RedisExcel') {
  if (-not $Quiet) { [System.Windows.Forms.MessageBox]::Show($text, $title) | Out-Null }
  Write-Host $text
}

# Host may already be "host:port" (or a connection string "host:port,password=..").
$target = $ServerHost.Trim()
if ($target -match '^\s*([^,:;]+)(?::(\d+))?') {
  $serverHost = $Matches[1]
  $serverPort = $Port
  if ($Matches[2]) { $serverPort = [int]$Matches[2] }
} else {
  $serverHost = $ServerHost
  $serverPort = $Port
}
$target = "$serverHost`:$serverPort"

function Connect-Redis([string]$h, [int]$p) {
  $client = New-Object System.Net.Sockets.TcpClient    # no timeout in the BCL ctor
  $connect = $client.ConnectAsync($h, $p)
  if (-not $connect.Wait($ConnectTimeoutMs)) {
    $client.Close()
    throw ("no TCP answer within " + [int]($ConnectTimeoutMs / 1000) + "s (unreachable host or firewall?)")
  }
  return $client
}

# A failed Task.Wait surfaces as an AggregateException whose own message says
# nothing ("Exception calling Wait with 1 argument(s)"); the innermost exception
# carries the reason the user needs ("... actively refused it").
function Get-ErrorMessage($record) {
  $ex = $record.Exception
  while ($ex -and $ex.InnerException) { $ex = $ex.InnerException }
  if ($ex) { return $ex.Message }
  return $record.Exception.Message
}

# RESP array of bulk strings: the only form that can carry a password with spaces.
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

function Test-Redis([string]$h, [int]$p) {
  try {
    $client = Connect-Redis $h $p
    $stream = $client.GetStream()
    $reader = New-Object System.IO.StreamReader($stream)
    # A server with requirepass answers -NOAUTH to a bare PING, so the AUTH must
    # happen first: otherwise a perfectly good install is reported as unreachable.
    if ($Password -ne '') {
      if ($User -ne '') { Send-Command $stream @('AUTH', $User, $Password) }
      else { Send-Command $stream @('AUTH', $Password) }
      $auth = $reader.ReadLine()
      if ($auth -notmatch '^\+OK') {
        Write-Host "[verify] AUTH failed: $auth"
        $client.Close()
        return $false
      }
    }
    # Real protocol PING, not just a socket connect.
    Send-Command $stream @('PING')
    $pong = $reader.ReadLine()
    $client.Close()
    return ($pong -match 'PONG')
  } catch {
    Write-Host ("[verify] no answer: " + (Get-ErrorMessage $_))
    return $false
  }
}

$target = "$serverHost`:$serverPort"
Write-Host "[verify] testing Redis at $target ..."
$ok = Test-Redis $serverHost $serverPort

if (-not $ok) {
  Show ("Could not connect to Redis at $target.`r`n`r`n" +
        "- Check that the server is running (docker ps, or the service).`r`n" +
        "- If it is on another host/port, adjust RedisExcel.json.") "RedisExcel - connection"
  Hold
  exit 1
}

Write-Host "[verify] Redis answered PONG at $target"

if (-not $OpenExcel) {
  Show ("All set!`r`n`r`nRedis answered at $target.`r`n" +
        "Open Excel, register the XLL (File > Options > Add-ins > Go... > Browse...) " +
        "and use, for example, =RedisUDFGet(`"key`").") "RedisExcel - installed"
  Hold
  exit 0
}

# ---- Open a sample sheet in Excel ------------------------------------------
$xll64 = Join-Path $AddInDir 'RedisExcel64-packed.xll'
$xll32 = Join-Path $AddInDir 'RedisExcel-packed.xll'
$xll = if (Test-Path $xll64) { $xll64 } elseif (Test-Path $xll32) { $xll32 } else { $null }
if (-not $xll) {
  Show ("Could not find the XLL in $AddInDir.") "RedisExcel"
  Hold
  exit 0
}

# With a password the cells must NOT pass an explicit host: the no-argument form
# makes the add-in use the configured server (RedisExcel.json carries the secret),
# so no password ever reaches a cell formula - and the status cell would otherwise
# read FAILED against a server that requires AUTH.
$hostArg = if ($Password -ne '') { '' } else { '("' + $target + '")' }
$exampleArg = if ($Password -ne '') { '' } else { ', "' + $target + '"' }

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $registered = $excel.RegisterXLL($xll)
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'RedisExcel'

  # Nice title block. English on purpose: the add-in is English-only inside
  # Excel (a localized volatile result would differ per machine), and a sample
  # sheet is the first thing a user sees.
  $ws.Range('A1').Value2 = 'RedisExcel installed successfully!'
  $ws.Range('A1').Font.Size = 16
  $ws.Range('A1').Font.Bold = $true
  $ws.Range('A2').Value2 = 'Redis connection test - ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  $ws.Range('A2').Font.Italic = $true

  $ws.Range('A4').Value2 = 'Server:'
  $ws.Range('B4').Value2 = $target
  $ws.Range('A5').Value2 = 'Status:'
  # ServerTime is a real round trip; if it errors the cell shows Error:.
  $ws.Range('B5').Formula = '=IF(ISERROR(RedisUDFServerTime' + $hostArg + '),"FAILED","OK - Redis answered")'
  $ws.Range('A6').Value2 = 'Server says:'
  $ws.Range('B6').Formula = '=RedisUDFServerTime' + $hostArg
  # Version and the XLL that answered: the first question of any support case,
  # and how a SECOND installed copy becomes obvious (each one reports its own
  # path in the xll= field).
  $ws.Range('A7').Value2 = 'Version:'
  $ws.Range('B7').Formula = '=RedisUDFVersion()'

  # The sample workbook is the USER'S OWN sheet, so a write here is fine and
  # useful: it shows a real round trip (Set then Get). This is deliberately
  # different from the installer's connection PROBE (test-connection.ps1), which
  # must never write to a server it does not own.
  $ws.Range('A9').Value2 = 'Read/write test (press F9 to refresh):'
  $ws.Range('A10').Value2 = 'Set:'
  $ws.Range('B10').Formula = '=RedisUDFSet("redisexcel:install_test","hello from the installer")'
  $ws.Range('A11').Value2 = 'Get:'
  $ws.Range('B11').Formula = '=RedisUDFGet("redisexcel:install_test")'

  $ws.Range('A13').Value2 = 'Example: =RedisUDFGet("price_btc"' + $exampleArg + ')'
  $ws.Range('A4:A7').Font.Bold = $true
  $ws.Range('A10:A11').Font.Bold = $true
  $ws.Range('A').ColumnWidth = 42
  $ws.Range('B').ColumnWidth = 40
  $ws.Range('B5').Font.Bold = $true
  $ws.Range('B7').Font.Italic = $true

  $excel.Application.Calculate() | Out-Null
  Write-Host "[verify] sample workbook opened (XLL registered: $registered)"
  Hold
  exit 0
} catch {
  Show ("Could not open Excel: " + $_.Exception.Message) "RedisExcel"
  Hold
  exit 0
}
