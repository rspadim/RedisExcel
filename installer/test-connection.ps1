<#
  test-connection.ps1 - quick Redis connection test used by the installer wizard.

  Speaks to the server through **StackExchange.Redis**, the same client the
  add-in uses, instead of a hand-written RESP exchange. That matters because the
  client already knows every case a raw PING does not: ACL users, TLS
  (`rediss://`), a Redis Cluster (MOVED/ASK), a Sentinel connection string and
  RESP3. A connection string accepted here is exactly one the add-in accepts.

  The DLL is loaded from the add-in folder (it always ships next to the XLL). If
  it is missing, the script falls back to the small RESP probe it used before,
  so a bare copy of this script still works.

  Exit codes are STABLE reason codes so a caller can show its own (localized)
  message instead of this script's text: 0 = PONG, 2 = unknown host (DNS),
  3 = timeout/unreachable, 4 = connection refused or other socket error,
  5 = AUTH failed, 6 = unexpected reply.

  Examples:
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:6379
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost localhost:6379 -PasswordFile C:\Temp\pw.txt
    powershell -ExecutionPolicy Bypass -File test-connection.ps1 -ServerHost "sentinel-a:26379,serviceName=mymaster" -Quiet
#>
[CmdletBinding()]
param(
  [string]$ServerHost = 'localhost',
  [int]$Port = 6379,
  [string]$User = '',
  [string]$Password = '',
  [string]$PasswordFile = '',
  # Suppress the console header (the wizard redirects this output to a file and
  # only reads the exit code).
  [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$ConnectTimeoutMs = 5000

# The file content is used verbatim (no Trim): a password may contain spaces.
if ($PasswordFile -ne '' -and (Test-Path -LiteralPath $PasswordFile)) {
  $Password = [System.IO.File]::ReadAllText($PasswordFile)
}

# Console header (see banner.ps1), unless the caller asked to stay quiet.
if (-not $Quiet) {
  try {
    $bannerPath = Join-Path $PSScriptRoot 'banner.ps1'
    if (Test-Path -LiteralPath $bannerPath) {
      . $bannerPath
      Write-RedisExcelBanner -Action ('Checking the Redis connection at ' + $ServerHost)
    }
  } catch { }
}

# --- exit-code classification ----------------------------------------------
# The client wraps every failure: PowerShell adds its own
# "Exceção ao chamar "Execute"..." layer, then StackExchange.Redis adds a
# "message timed out in the backlog" layer, and only the innermost message names
# the real reason ("NOAUTH Authentication required", "UnableToConnect on
# host:port", "MISCONF ..."). Both the classification below and the line printed
# for the user need that inner text, so it is extracted once.
function Get-InnermostError {
  param($Record)
  $ex = $Record.Exception
  while ($ex -and $ex.InnerException) { $ex = $ex.InnerException }
  return $ex
}

# Every message in the chain, so a marker is found wherever the client put it.
function Get-ErrorText {
  param($Record)
  $text = ''
  $ex = $Record.Exception
  while ($ex) {
    $text = $text + ' ' + $ex.Message
    $ex = $ex.InnerException
  }
  return $text
}

# The user-facing reason, with the wrapper noise removed. PowerShell's own
# "Exception calling X with N argument(s)" carries no information for a human, so
# the innermost message is preferred and trimmed of stray quotes/whitespace.
function Get-FriendlyReason {
  param($Record)

  $inner = Get-InnermostError $Record
  $reason = $inner.Message

  # Unwrap the client's own "The message timed out in the backlog ... (5000ms) -
  # Last connection error: <real reason>" form when the real reason is inside it.
  if ($reason -match 'Last connection error:\s*(.+)$') {
    $reason = $Matches[1]
  }
  # Drop a leading "no answer: " style prefix the client may add.
  $reason = $reason -replace '^\s*no answer:\s*', ''
  # Collapse the long multi-line diagnostic the client emits after a comma.
  $reason = ($reason -replace '\s+', ' ').Trim()
  # The client appends a very long internal state dump after the first sentence;
  # keep the sentence a person can act on.
  $cut = $reason.IndexOf(' on ')
  if ($cut -gt 0 -and $reason.Length -gt 120) { $reason = $reason.Substring(0, $cut) }

  if ($reason.Length -gt 200) { $reason = $reason.Substring(0, 197) + '...' }
  return $reason
}

function Get-FailureCode {
  param($Record)

  $text = Get-ErrorText $Record

  # Authentication: the client reports AuthenticationFailure and, in the inner
  # exception, the server's own NOAUTH/WRONGPASS/ERR reply.
  if ($text -match 'AuthenticationFailure|NOAUTH|WRONGPASS|invalid password|NOPERM') { return 5 }

  # A socket-level cause, when one did survive (the raw RESP fallback path).
  $ex = $Record.Exception
  while ($ex) {
    $se = $ex -as [System.Net.Sockets.SocketException]
    if ($se) {
      if ($se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound -or
          $se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::NoData) { return 2 }
      if ($se.SocketErrorCode -eq [System.Net.Sockets.SocketError]::TimedOut) { return 3 }
      return 4
    }
    $ex = $ex.InnerException
  }

  # Name resolution failures, as the client spells them.
  if ($text -match 'No such host|host (name )?(could not be resolved|not known|unknown)|Name or service not known') { return 2 }

  # A connect timeout (unreachable host / firewall) versus a refusal. The client
  # does not always distinguish them: an unopened local port usually surfaces as
  # a plain UnableToConnect, the same shape as an unreachable host. Both land on
  # the same advice ("is the server running on that port / reachable?"), so a
  # code 3 here is correct even when the port is closed rather than filtered.
  if ($text -match 'ConnectTimeout|Connection refused|actively refused|No connection could be made') { return 3 }
  if ($text -match 'UnableToConnect') { return 3 }

  return 6
}

# --- the client path --------------------------------------------------------
# Build the same connection string the add-in would use for these fields, then
# let the client parse it, so a sentinel/cluster/TLS string behaves identically
# in the probe and in Excel.
function Connect-WithClient {
  param([string]$DllPath)

  Add-Type -Path $DllPath

  $endpoint = $ServerHost.Trim()
  if ($endpoint -notmatch ',\s*(ssl|password|user|serviceName|abortConnect|defaultDatabase)') {
    # A bare host (or host:port): add the port the wizard collected.
    if ($endpoint -notmatch ':\d+$' -and $Port -gt 0) { $endpoint = $endpoint + ':' + $Port }
  }

  # Resolve the host FIRST: the client reports a name-resolution failure as a
  # generic UnableToConnect, indistinguishable from an unreachable server, and
  # the two need different advice ("check the address" vs "is the server up?").
  # A sentinel/cluster string is left alone: its first entry is a real server and
  # the client resolves the rest itself.
  $hostPart = $endpoint
  if ($hostPart -match '^([^,:;]+)') { $hostPart = $Matches[1] }
  $hostPart = $hostPart -replace '^\w+://', ''
  if ($hostPart -and $hostPart -notmatch '^\d+\.\d+\.\d+\.\d+$' -and $hostPart -notmatch '^[0-9a-fA-F:]+$') {
    [void][System.Net.Dns]::GetHostAddresses($hostPart)
  }

  $options = [StackExchange.Redis.ConfigurationOptions]::Parse($endpoint)
  $options.AbortOnConnectFail = $false
  $options.ConnectTimeout = $ConnectTimeoutMs
  $options.SyncTimeout = $ConnectTimeoutMs
  # The wizard's own fields win over anything embedded in the string.
  if ($User -ne '') { $options.User = $User }
  if ($Password -ne '') { $options.Password = $Password }

  $mux = [StackExchange.Redis.ConnectionMultiplexer]::Connect($options)
  try {
    $db = $mux.GetDatabase()

    # Connection detection: ANY reply proves the whole path (TCP, TLS, AUTH,
    # database selection) - even an error reply. Only "no reply" means the
    # address is wrong. So the verdict is "the server answered something", and
    # the only real failures are auth rejection and non-permission on EVERY
    # probe tried.
    #
    # A ladder, because each command belongs to a different ACL category and a
    # restricted user (exactly the kind a read-only feed grants) may be denied
    # one while being allowed the next:
    #   PUBSUB CHANNELS  @pubsub      - what the add-in actually does (SUB/PSUB)
    #   DBSIZE           @keyspace    - what a GET user always has
    #   PING             @connection  - allowed to nearly everyone
    #
    # PING alone is a BAD verdict: a Redis with `stop-writes-on-bgsave-error`
    # answers it with "-MISCONF ..." even though every read works, so it
    # rejects a server the add-in uses happily. DBSIZE alone is bad too: an
    # @read/@pubsub-only user gets -NOPERM. And PUBLISH is never an option -
    # the installer must not write to, or publish into, a server it does not
    # own.
    $connected = $false
    $authProblem = $null
    $probingError = $null
    foreach ($command in @("PUBSUB CHANNELS", "DBSIZE", "PING")) {
      try {
        if ($command -eq "PUBSUB CHANNELS") {
          # Bounded by the same timeout; a channel list is read-only metadata.
          [void]$db.Execute("PUBSUB", "CHANNELS")
        } else {
          [void]$db.Execute($command)
        }
        $connected = $true
        break
      } catch {
        # Keep the ORIGINAL ErrorRecord: its exception chain still carries the
        # client's real reason (NOAUTH, UnableToConnect, ...). Rebuilding an
        # exception from the concatenated text would lose it and print the
        # PowerShell wrapper instead.
        $probingError = $_

        $text = Get-ErrorText $_

        # Credentials are wrong for the whole server, not for one command: no
        # point trying the rest.
        if ($text -match 'NOAUTH|WRONGPASS|AuthenticationFailure|invalid password') {
          $authProblem = Get-FriendlyReason $_
          break
        }
        # Denied for this command only (-NOPERM): try the next one.
        if ($text -match 'NOPERM|no permissions') { continue }
        $authProblem = Get-FriendlyReason $_
        continue
      }
    }

    if (-not $connected) {
      # Re-throw the ORIGINAL record so the caller's classifier sees the same
      # exception chain a direct connect failure would produce.
      if ($probingError) { throw $probingError }
      throw (New-Object System.Exception([string]$authProblem))
    }
  } finally {
    $mux.Dispose()
  }
}

$clientDll = $null
foreach ($candidate in @(
    (Join-Path (Split-Path -Parent $PSScriptRoot) 'StackExchange.Redis.dll'),
    (Join-Path $PSScriptRoot 'StackExchange.Redis.dll'))) {
  if ($candidate -and (Test-Path -LiteralPath $candidate)) { $clientDll = $candidate; break }
}

if ($clientDll) {
  try {
    Connect-WithClient -DllPath $clientDll
    Write-Host 'PONG'
    exit 0
  } catch {
    $code = Get-FailureCode $_
    $reason = Get-FriendlyReason $_
    Write-Host ('no answer: ' + $reason)
    exit $code
  }
}

# --- fallback: raw RESP -----------------------------------------------------
# Only reached when the client DLL is not next to the script (a lone copy of
# this file). Covers the common standalone case: AUTH (optional) + PING. ACL
# users, TLS and cluster are NOT handled here - the wizard always ships the DLL.
Write-Host 'note: StackExchange.Redis.dll not found next to this script; using the raw RESP probe.'

$hostName = $ServerHost
$port = $Port
if ($ServerHost -match '^\s*([^,:;]+)(?::(\d+))?') {
  $hostName = $Matches[1]
  if ($Matches[2]) { $port = [int]$Matches[2] }
}

# RESP array of bulk strings: the only form that can carry a password with
# spaces and binary bytes.
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

$client = $null
try {
  $client = New-Object System.Net.Sockets.TcpClient    # no timeout in the BCL ctor
  $connect = $client.ConnectAsync($hostName, $port)
  if (-not $connect.Wait($ConnectTimeoutMs)) {
    $client.Close()
    Write-Host ('no answer within ' + [int]($ConnectTimeoutMs / 1000) + 's')
    exit 3
  }
  $stream = $client.GetStream()
  $reader = New-Object System.IO.StreamReader($stream)

  if ($Password -ne '') {
    if ($User -ne '') { Send-Command $stream @('AUTH', $User, $Password) }
    else { Send-Command $stream @('AUTH', $Password) }
    $auth = $reader.ReadLine()
    if ($auth -notmatch '^\+OK') {
      Write-Host "AUTH failed: $auth"
      $client.Close()
      exit 5
    }
  }

  Send-Command $stream @('PING')
  $pong = $reader.ReadLine()
  $client.Close()
  if ($pong -match 'PONG') { Write-Host 'PONG'; exit 0 }
  Write-Host "unexpected reply: $pong"
  exit 6
} catch {
  $code = Get-FailureCode $_
  $ex = Get-InnermostError $_
  Write-Host ('no answer: ' + $ex.Message)
  exit $code
} finally {
  if ($client) { try { $client.Close() } catch { } }
}
