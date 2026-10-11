<#
.SYNOPSIS
End-to-end test for the RedisExcel add-in using a real (hidden) Excel instance.

.DESCRIPTION
Builds the test workbook and asserts UDF/RTD results against a running Redis
server, including the v1.1.0 Pub/Sub regression scenario:

  1. Loads the packed XLL into Excel (RegisterXLL).
  2. Creates the UDF/RTD test sheets and saves test\RedisExcel.Test.xlsx
     (only for local hosts without -AsyncWrites; remote hosts and async runs
     are saved to %TEMP% so the host never lands in the repository and the
     async scratch cells stay out of the committed sample).
  3. Asserts values (SET/GET/EXISTS/TTL/JSON/HASH, RTD GET/HGET/HGETALL/SUB/PSUB).
  4. Opens a COPY of the workbook, publishes messages, closes the copy and
     verifies the original workbook keeps receiving (the reported bug).
  5. Kills the Pub/Sub connections server-side (CLIENT KILL TYPE pubsub) and
     verifies the subscriptions recover automatically. This step only runs
     against local hosts; use -SkipClientKill to force it off.
  6. Kills its own hidden Excel mid-stream and verifies the Redis server drops
     every subscription (a crashed add-in leaves no dangling state; skipped
     with -KeepExcelOpen).

TESTS USE ONLY THE test.redisexcel.* CHANNELS AND test.redisexcel.* KEYS.

NOTES
  - Antivirus/EDR software may block Excel COM automation. If Excel is blocked
    the script fails with COM errors (sometimes misleading ones). Add an
    exception for EXCEL.EXE and for this script before running.
  - Excel is busy while recalculating/starting the RTD server; every COM call
    here retries RPC_E_CALL_REJECTED (0x80010001) / RPC_E_SERVERCALL_RETRYLATER
    (0x8001010A).
  - This script uses its own hidden Excel instance; it never attaches to an
    Excel window the user already has open.

.PARAMETER RepoRoot
Repository root. Defaults to the parent folder of this script.

.PARAMETER RedisHost
Redis connection string used in the sheet formulas. Default localhost:6379.

.PARAMETER KeyPrefix
Prefix for every key and channel used by the test. Default test.redisexcel.

.PARAMETER RealChannel
Optional real (production) channel name to also subscribe with RTD SUB and
assert that live data arrives. Read-only.

.PARAMETER RealPattern
Optional real channel pattern to also subscribe with RTD PSUB and assert that
live data arrives. Read-only.

.PARAMETER SkipClientKill
Never run CLIENT KILL TYPE pubsub (automatically skipped for non-local hosts).

.PARAMETER RedisCli
Command used to publish and to run CLIENT KILL. Defaults to redis-cli on PATH,
or "docker exec redisexcel-test redis-cli" when that container is running.

.PARAMETER SyncWrite
Write delivery mode for this run: sync, fireforget (default) or
fireforget-all. The value is written to %USERPROFILE%\RedisExcel.json before
Excel starts; the user's own config file is restored afterwards.

.PARAMETER AsyncWrites
Dispatch writes asynchronously for this run (same config handling as
SyncWrite). The cell shows Excel's pending marker first, then the settled
reply; the checks wait for the settled text.

.PARAMETER KeepExcelOpen
Do not quit Excel at the end (useful for debugging).

.EXAMPLE
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1

.EXAMPLE
powershell -ExecutionPolicy Bypass -File test\Run-ExcelE2E.ps1 `
    -RedisHost redis.example.com -RealChannel test.redisexcel.real -SkipClientKill
#>
#Requires -Version 5.1
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$RedisHost = 'localhost:6379',
    [string]$KeyPrefix = 'test.redisexcel',
    [string]$RealChannel = $null,
    [string]$RealPattern = $null,
    [switch]$SkipClientKill,
    [string]$RedisCli = $null,
    [ValidateSet('sync', 'fireforget', 'fireforget-all')]
    [string]$SyncWrite = 'fireforget',
    [switch]$AsyncWrites,
    [switch]$KeepExcelOpen
)

$ErrorActionPreference = 'Stop'

# ValidateSet is case-insensitive; normalize so RedisExcel.json always carries
# the exact contract spelling.
$SyncWrite = $SyncWrite.ToLowerInvariant()

$script:Failures = 0
$script:RedisExe = $null
$script:RedisPrefix = @()
$script:RedisArgs = @()
# Attempt count of the last Invoke-ExcelAction call, read by the Set-*/Get-*
# helpers to report the "succeeded after N retries" diagnostic.
$script:ExcelActionAttempts = 0
$script:Excel = $null
$script:Workbook = $null
$script:ExcelPidsBefore = @()
# Set when the "Excel dies mid-stream" scenario intentionally kills this run's
# hidden Excel: the finally must then skip the COM close/quit calls and the
# wait-for-exit loop (the COM proxy is dead; the process is already gone).
$script:ExcelKilled = $false

# ------------------------------------------------------------- write mode ----
# v1.3.0 SyncWrite/AsyncWrites (written to RedisExcel.json before Excel starts;
# see the config block in the setup section):
#   sync           - historical behavior: every write returns its real reply.
#   fireforget     - reply-agnostic writes return 'OK (fire and forget)'; reply-dependent
#                    writes (Del, Incr, IncrBy, Expire, SetAdd, SetRemove,
#                    HashDel, ListPopRight/Left) keep their real replies.
#   fireforget-all - every write is fire-and-forget: reply-agnostic writes keep
#                    'OK (fire and forget)', reply-dependent ones return
#                    'OK (fire and forget: all)'.
# AsyncWrites changes only when the reply arrives (the cell shows Excel's
# pending marker first); every expectation below waits for the settled text.
$script:IsFireForget = ($SyncWrite -ne 'sync')
$script:IsFireForgetAll = ($SyncWrite -eq 'fireforget-all')
# Reply-agnostic write ack (Set, SetJSON, HashSet, SetEx, Rename, ...).
$script:WriteAck = if ($script:IsFireForget) { 'OK (fire and forget)' } else { 'OK' }
# Reply-dependent writes: a numeric reply, or the marker when everything is FF.
$script:IntReplyPattern = if ($script:IsFireForgetAll) { '^OK \(fire and forget: all\)$' } else { '^\d+$' }
$script:ZeroReplyExpected = if ($script:IsFireForgetAll) { 'OK (fire and forget: all)' } else { '0' }
# ListPushRight/Left are reply-agnostic (an integer reply only in sync mode).
$script:ListPushReplyPattern = if ($script:IsFireForget) { '^OK \(fire and forget\)$' } else { '^\d+$' }
# ChannelPublish reports the readers count only in sync mode.
$script:ChannelPublishExpected = if ($script:IsFireForget) { 'OK (fire and forget)' } else { '1 reader' }
# ListPopRight/Left on a missing list return "" only when the reply is awaited.
$script:ListPopEmptyExpected = if ($script:IsFireForgetAll) { 'not empty' } else { 'empty' }
# Rename of a missing key reports the error only when the reply is awaited.
$script:RenameMissingPattern = if ($script:IsFireForget) { '^OK \(fire and forget\)$' } else { '^Error' }

function Check([bool]$Condition, [string]$Label) {
    if ($Condition) { Write-Host ("PASS " + $Label) -ForegroundColor Green }
    else { Write-Host ("FAIL " + $Label) -ForegroundColor Red; $script:Failures++ }
}

function Get-RedisEndpoint([string]$ConnectionString) {
    $first = ($ConnectionString -split ',')[0].Trim()
    if ($first -match '^(?<host>[^:]+)(:(?<port>\d+))?$') {
        $port = if ($Matches['port']) { [int]$Matches['port'] } else { 6379 }
        return @{ Host = $Matches['host']; Port = $port }
    }
    throw "Cannot parse the Redis endpoint from '$ConnectionString'"
}

function Invoke-RedisCli([string[]]$Arguments) {
    return (& $script:RedisExe @($script:RedisPrefix + $script:RedisArgs + $Arguments))
}

# Reads the numeric part of a redis-cli PUBSUB reply: NUMSUB answers
# "<channel>\n<count>" (the count is the last line), NUMPAT just "<count>".
# Returns -1 when the reply is empty, missing or not a number.
function Get-RedisPubsubCount([string[]]$Arguments) {
    $lines = @((Invoke-RedisCli $Arguments | Out-String).Trim() -split '\r?\n')
    $lastLine = ''
    if ($lines.Count -gt 0 -and $lines[-1]) { $lastLine = $lines[-1].Trim() }
    # TryParse zeroes the ref value on failure, so test its return explicitly:
    # a garbage reply must yield -1, never the "clean" value 0.
    $value = 0
    if ([int]::TryParse($lastLine, [ref]$value)) { return $value }
    return -1
}

# Trimmed text of a single-value redis-cli reply (GET/LLEN/PUBLISH/... probes).
# The diagnostic labels around each call site are unchanged.
function Get-RedisValue([string[]]$Arguments) {
    return ((Invoke-RedisCli $Arguments | Out-String).Trim())
}

function Resolve-RedisCli {
    $endpoint = Get-RedisEndpoint $RedisHost
    $script:RedisArgs = @('-h', $endpoint.Host, '-p', $endpoint.Port)
    if ($RedisCli) {
        $parts = $RedisCli -split '\s+'
        $script:RedisExe = $parts[0]
        $script:RedisPrefix = @($parts | Select-Object -Skip 1)
        $script:RedisArgs = @()
        Write-Host ("Custom Redis CLI: -h/-p not added; make sure it targets {0}" -f $RedisHost) -ForegroundColor DarkYellow
        return
    }
    if (Get-Command redis-cli -ErrorAction SilentlyContinue) {
        $script:RedisExe = 'redis-cli'
        return
    }
    if (Get-Command docker -ErrorAction SilentlyContinue) {
        $container = ((& docker ps --filter name=redisexcel-test --format "{{.Names}}" 2>$null) -join '').Trim()
        if ($container -eq 'redisexcel-test') {
            $script:RedisExe = 'docker'
            $script:RedisPrefix = @('exec', 'redisexcel-test', 'redis-cli')
            return
        }
    }
    throw "redis-cli not found and container 'redisexcel-test' is not running. Pass -RedisCli (e.g. 'docker exec my-redis redis-cli')."
}

function Test-ComBusyError($ErrorRecord) {
    $hr = $ErrorRecord.Exception.HResult
    # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER and the generic Excel
    # rejection 0x800A03EC (seen from CalculateFull/CalculateBeforeSave while
    # Excel is busy with pending async work) are all transient here.
    return ($hr -eq -2147418111 -or $hr -eq -2147417846 -or $hr -eq -2146827284)
}

# Some security products intermittently break Excel COM property sets with a
# bogus InvalidCastException (e.g. "cannot convert Int32 to String"), and under
# load an interop call can come back null mid-expression ("You cannot call a
# method on a null-valued expression"). Those failures are transient; retry
# them like a busy Excel (a persistent null still surfaces after the retries).
function Test-RetryableError($ErrorRecord) {
    if ($ErrorRecord.Exception -is [System.InvalidCastException]) { return $true }
    if ($ErrorRecord.FullyQualifiedErrorId -like 'InvokeMethodOnNull*') { return $true }
    return (Test-ComBusyError $ErrorRecord)
}

# Reports the registered antivirus products through the Windows Security
# Center. Security software is the top cause of intermittent Excel COM
# failures (see AGENTS.md), so the status is printed before Excel starts: a
# failed or flaky run is then explained immediately. Best effort - where
# SecurityCenter2 is unavailable (some domain policies) that is reported too,
# never fatal.
function Get-AntivirusStatus {
    try {
        return @(Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop | ForEach-Object {
            $state = [int]$_.productState
            # The vendor-specific enabled/snoozed code lives in the 0x?000
            # nibble (0 = off; 1..3 = on/snoozed variants; e.g. Kaspersky
            # reports 0x42000 while active). The raw state is printed too.
            $level = ($state -shr 12) -band 0xF
            [pscustomobject]@{
                Name = [string]$_.displayName
                Active = ($level -ge 1)
                StateHex = ('0x{0:X}' -f $state)
            }
        })
    }
    catch {
        return @([pscustomobject]@{
            Name = "unavailable (" + $_.Exception.Message + ")"
            Active = $false
            StateHex = ''
        })
    }
}

function Invoke-ExcelAction([scriptblock]$Action, [int]$Retries = 40) {
    for ($attempt = 0; $attempt -lt $Retries; $attempt++) {
        try {
            $script:ExcelActionAttempts = $attempt
            return (& $Action)
        }
        catch {
            if (-not (Test-RetryableError $_)) {
                $script:ExcelActionAttempts = $attempt
                throw
            }
            Start-Sleep -Milliseconds ([Math]::Min(150 * ($attempt + 1), 3000))
        }
    }
    $script:ExcelActionAttempts = $Retries
    return (& $Action)
}

# Single bounded poll used by the Wait-* family: runs $Probe every $IntervalMs
# until it returns true or $TimeoutSeconds elapse.
function Wait-Until([scriptblock]$Probe, [int]$TimeoutSeconds, [int]$IntervalMs = 120) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (& $Probe) { return $true }
        Start-Sleep -Milliseconds $IntervalMs
    }
    return $false
}

function Get-CellAddress([int]$Row, [int]$Col) {
    $letter = ''
    while ($Col -gt 0) {
        $rem = ($Col - 1) % 26
        $letter = [char](65 + $rem) + $letter
        $Col = [int](($Col - 1) / 26)
    }
    return "$letter$Row"
}

function Set-Cell($Sheet, [int]$Row, [int]$Col, $Value) {
    $address = Get-CellAddress $Row $Col
    # Values are written as strings. Mixing value types through the same
    # PowerShell call site can hit a COM member-cache bug (intermittent
    # "cannot convert Int32 to String") that security software makes worse.
    # The unit tests cover numeric JSON conversion; the workbook only needs text.
    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    try {
        Invoke-ExcelAction { $Sheet.Range($address).Value2 = $text } | Out-Null
    }
    catch {
        Write-Host ("Set-Cell[{0}] failed after {1} attempts: {2}" -f $address, $script:ExcelActionAttempts, $_.Exception.Message) -ForegroundColor Red
        throw
    }
    if ($script:ExcelActionAttempts -gt 0) {
        Write-Host ("      Set-Cell[{0}] succeeded after {1} retries (transient COM failure)" -f $address, $script:ExcelActionAttempts) -ForegroundColor DarkYellow
    }
}

function Set-Formula($Sheet, [string]$Address, [string]$Formula) {
    try {
        Invoke-ExcelAction { $Sheet.Range($Address).Formula = $Formula } | Out-Null
    }
    catch {
        Write-Host ("Set-Formula[{0}] failed after {1} attempts: {2}" -f $Address, $script:ExcelActionAttempts, $_.Exception.Message) -ForegroundColor Red
        throw
    }
    if ($script:ExcelActionAttempts -gt 0) {
        Write-Host ("      Set-Formula[{0}] succeeded after {1} retries (transient COM failure)" -f $Address, $script:ExcelActionAttempts) -ForegroundColor DarkYellow
    }
}

function Get-CellText($Sheet, [string]$Address) {
    return [string](Invoke-ExcelAction { [string]$Sheet.Range($Address).Text })
}

function Wait-CellText($Sheet, [string]$Address, [string]$Expected, [int]$TimeoutSeconds = 20) {
    $last = @('')
    $ok = Wait-Until { $last[0] = Get-CellText $Sheet $Address; return ($last[0] -eq $Expected) } $TimeoutSeconds
    if ($ok) { return $true }
    Write-Host ("      {0} = '{1}' (expected '{2}')" -f $Address, $last[0], $Expected) -ForegroundColor DarkGray
    return $false
}

function Wait-CellNotEmpty($Sheet, [string]$Address, [int]$TimeoutSeconds = 20) {
    $last = @('')
    $ok = Wait-Until {
        $last[0] = Get-CellText $Sheet $Address
        return (-not [string]::IsNullOrWhiteSpace($last[0]) -and
            $last[0] -ne '(ConnectData)' -and
            -not $last[0].StartsWith('#'))
    } $TimeoutSeconds
    if ($ok) { return $true }
    Write-Host ("      {0} = '{1}' (empty/placeholder after {2}s)" -f $Address, $last[0], $TimeoutSeconds) -ForegroundColor DarkGray
    return $false
}

function Wait-CellRegex($Sheet, [string]$Address, [string]$Pattern, [int]$TimeoutSeconds = 20) {
    $last = @('')
    $ok = Wait-Until { $last[0] = Get-CellText $Sheet $Address; return ($last[0] -match $Pattern) } $TimeoutSeconds
    if ($ok) { return $true }
    Write-Host ("      {0} = '{1}' (expected /{2}/)" -f $Address, $last[0], $Pattern) -ForegroundColor DarkGray
    return $false
}

# #N/A is Excel's N/A error: the display text is locale-dependent (#N/A,
# #N/D, #NV, ...), so a pending async cell is detected through
# WorksheetFunction.IsNA instead of matching the text. The Range is passed
# directly: reading Value2 first loses the error type (it surfaces as Int32),
# and IsNA on that number is False. A failed check (retries exhausted or a
# non-retryable COM error) returns $false.
function Test-CellIsNA($Sheet, [string]$Address) {
    try { return [bool](Invoke-ExcelAction { [bool]$script:Excel.WorksheetFunction.IsNA($Sheet.Range($Address)) }) }
    catch { return $false }
}

function Wait-CellIsNA($Sheet, [string]$Address, [int]$TimeoutSeconds = 5) {
    $ok = Wait-Until { Test-CellIsNA $Sheet $Address } $TimeoutSeconds
    if ($ok) { return $true }
    Write-Host ("      {0} = '{1}' (expected the locale-independent #N/A pending error)" -f $Address, (Get-CellText $Sheet $Address)) -ForegroundColor DarkGray
    return $false
}

function Wait-CellNumberMin($Sheet, [string]$Address, [double]$Min, [int]$TimeoutSeconds = 20) {
    $last = @('')
    $ok = Wait-Until {
        $last[0] = (Get-CellText $Sheet $Address).Replace(',', '.')
        $value = 0.0
        return ([double]::TryParse($last[0], [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value) -and $value -ge $Min)
    } $TimeoutSeconds
    if ($ok) { return $true }
    Write-Host ("      {0} = '{1}' (expected >= {2})" -f $Address, $last[0], $Min) -ForegroundColor DarkGray
    return $false
}

# Polls a Redis command reply (e.g. GET/LLEN) until it equals the expected
# text; prints the last observed reply on timeout so failures are diagnosable.
function Wait-RedisValue([string[]]$Arguments, [string]$Expected, [int]$TimeoutSeconds = 5) {
    $last = @('')
    $ok = Wait-Until { $last[0] = Get-RedisValue $Arguments; return ($last[0] -eq $Expected) } $TimeoutSeconds
    if ($ok) { return $true }
    Write-Host ("      Redis {0} = '{1}' (expected '{2}')" -f ($Arguments -join ' '), $last[0], $Expected) -ForegroundColor DarkGray
    return $false
}

# A workbook may (re)register its RTD topics asynchronously; a single publish
# can be lost before that happens, so repeat it until the cell shows the value.
function Publish-Until-Cell($Channel, $Message, $Sheet, [string]$Address, [string]$Expected, [int]$TimeoutSeconds = 30) {
    $last = @('')
    $ok = Wait-Until {
        $last[0] = Get-RedisValue @('PUBLISH', $Channel, $Message)
        return (Wait-CellText $Sheet $Address $Expected 3)
    } $TimeoutSeconds 250
    if ($ok) { return $true }
    Write-Host ("      PUBLISH {0} '{1}' last reply: '{2}' (cell {3} never showed '{4}')" -f $Channel, $Message, $last[0], $Address, $Expected) -ForegroundColor DarkGray
    return $false
}

# Removes local machine paths and personal metadata from a saved workbook:
# Excel stores the save folder in xl/workbook.xml (x15ac:absPath, e.g. the
# user's TEMP path) and the author in docProps/core.xml. The committed sample
# must not carry either (see the no-private-data rule in AGENTS.md).
function Remove-WorkbookMetadata([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $rules = @{
            'xl/workbook.xml' = @(
                @{ Pattern = '(?s)<\w+:\w*absPath[^>]*/>'; Replacement = '' }
            )
            'docProps/core.xml' = @(
                @{ Pattern = '(?s)<dc:creator>.*?</dc:creator>'; Replacement = '<dc:creator>RedisExcel</dc:creator>' }
                @{ Pattern = '(?s)<cp:lastModifiedBy>.*?</cp:lastModifiedBy>'; Replacement = '<cp:lastModifiedBy>RedisExcel</cp:lastModifiedBy>' }
            )
            'docProps/app.xml' = @(
                @{ Pattern = '(?s)<Company>.*?</Company>'; Replacement = '<Company></Company>' }
                @{ Pattern = '(?s)<Manager>.*?</Manager>'; Replacement = '<Manager></Manager>' }
            )
        }
        foreach ($entryName in $rules.Keys) {
            $entry = $zip.GetEntry($entryName)
            if (-not $entry) { continue }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $clean = $text
            foreach ($rule in $rules[$entryName]) {
                $clean = [regex]::Replace($clean, $rule.Pattern, $rule.Replacement)
            }
            if ($clean -eq $text) { continue }
            # Update mode cannot rewrite an entry in place: delete and recreate
            # only the entries whose content changed.
            $entry.Delete()
            $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($entryName).Open())
            try { $writer.Write($clean) } finally { $writer.Dispose() }
        }
    }
    finally {
        $zip.Dispose()
    }
}

# ---------------------------------------------------------------- setup ----

Resolve-RedisCli
Write-Host ("Redis CLI : {0} {1}" -f $script:RedisExe, ($script:RedisPrefix -join ' '))
Write-Host ("Redis host: {0}" -f $RedisHost)
Write-Host ("Key prefix: {0}.*" -f $KeyPrefix)

if ((Get-RedisValue @('PING')) -ne 'PONG') {
    throw "Redis is not responding at $RedisHost"
}

$isLocalHost = $RedisHost -match '^(localhost|127\.0\.0\.1)(:\d+)?$'
$allowClientKill = (-not $SkipClientKill) -and $isLocalHost -and [string]::IsNullOrWhiteSpace($RedisCli)

# The add-in reads SyncWrite/AsyncWrites/SkipRepeatedMessages from
# RedisExcel.json once per Excel process; the copy in the user profile wins over
# the Excel folder and C:\Windows. Write a minimal config BEFORE Excel starts and
# restore the user's own file (or remove ours) in the finally block - never lose
# the config. SkipRepeatedMessages is pinned true so the duplicate-suppression
# assertions (e.g. B44 PublishIfChanged) always exercise the suppression branch
# instead of depending on a machine-local config.
$configPath = Join-Path $env:USERPROFILE 'RedisExcel.json'
$configBackup = $null
if (Test-Path -LiteralPath $configPath) {
    $configBackup = Join-Path $env:TEMP ("RedisExcel.json.bak-" + [Guid]::NewGuid().ToString('N'))
    Copy-Item -LiteralPath $configPath -Destination $configBackup -Force
}
$asyncJson = if ($AsyncWrites.IsPresent) { 'true' } else { 'false' }
$configJson = '{"SyncWrite":"' + $SyncWrite + '","AsyncWrites":' + $asyncJson + ',"SkipRepeatedMessages":true}'
[System.IO.File]::WriteAllText($configPath, $configJson)
Write-Host ("Sync write: {0} (AsyncWrites: {1})" -f $SyncWrite, $AsyncWrites.IsPresent)
Write-Host ("Config    : {0} -> {1}" -f $configPath, $configJson) -ForegroundColor DarkGray

# AV/EDR visibility: report the registered antivirus products before touching
# Excel, so a flaky run in a protected environment is identified immediately.
$avProducts = Get-AntivirusStatus
if ($avProducts.Count -eq 0) {
    Write-Host 'ENV antivirus: none reported by Windows Security Center' -ForegroundColor DarkGray
}
foreach ($av in $avProducts) {
    if ($av.Active) {
        Write-Host ("WARNING antivirus: {0} appears active ({1}) - Excel COM automation may fail intermittently; add an EXCEL.EXE exception when possible." -f $av.Name, $av.StateHex) -ForegroundColor Yellow
    }
    else {
        Write-Host ("ENV antivirus: {0} is not active ({1})" -f $av.Name, $av.StateHex) -ForegroundColor DarkGray
    }
}

$udf = $null
$rtd = $null

try {
    # Remember which EXCEL.EXE processes already existed so the cleanup waits
    # only for THIS run's hidden instance (never the user's own Excel).
    $script:ExcelPidsBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

    $script:Excel = New-Object -ComObject Excel.Application
    $script:Excel.Visible = $false
    $script:Excel.DisplayAlerts = $false
    $script:Excel.ScreenUpdating = $false

    $is32Bit = $script:Excel.Path -match '\(x86\)'
    $xllName = if ($is32Bit) { 'RedisExcel-packed.xll' } else { 'RedisExcel64-packed.xll' }
    $xll = Join-Path $RepoRoot "bin\Release\net48\publish\$xllName"
    if (-not (Test-Path $xll)) {
        throw "XLL not found: $xll. Build first: dotnet build RedisExcel.sln -c Release"
    }
    Check ([bool](Invoke-ExcelAction { $script:Excel.RegisterXLL($xll) })) ("RegisterXLL " + $xllName)

    $script:Workbook = Invoke-ExcelAction { $script:Excel.Workbooks.Add() }

    # Warm-up: security software may intermittently block the first COM property
    # writes after the XLL is loaded. Retry until a string and an int write succeed
    # (or give up with a clear message).
    $warmupSheet = $script:Workbook.Worksheets.Item(1)
    $warmupDeadline = (Get-Date).AddSeconds(90)
    while ($true) {
        try {
            $warmupSheet.Range('A1').Value2 = 'warmup'
            $warmupSheet.Range('A2').Value2 = 1
            break
        }
        catch {
            if ((Get-Date) -gt $warmupDeadline) {
                throw "Excel COM writes are being blocked (antivirus/EDR?): $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds 500
        }
    }
    $warmupSheet.Range('A1').Clear() | Out-Null
    $warmupSheet.Range('A2').Clear() | Out-Null

    # ------------------------------------------------------------ UDF sheet ----
    $udf = $script:Workbook.Worksheets.Item(1)
    $udf.Name = 'UDF'
    Set-Cell $udf 1 1 ("RedisExcel UDF tests (host " + $RedisHost + ", prefix " + $KeyPrefix + ")")
    Set-Cell $udf 3 1 'Function'
    Set-Cell $udf 3 2 'Live result'
    Set-Cell $udf 3 3 'Expected'
    Set-Cell $udf 2 6 1; Set-Cell $udf 2 7 'linha1'
    Set-Cell $udf 3 6 2; Set-Cell $udf 3 7 'linha2'
    Invoke-ExcelAction { $udf.Range('H2').Value2 = 67000.5 } | Out-Null

    $h = $RedisHost
    $kp = $KeyPrefix
    $udfItems = @(
        @{ Row = 4;  Func = 'Set';            Fx = '=RedisUDFSet("{1}.key","hello_from_udf","{0}")' -f $h, $kp;                                             Expected = 'OK' },
        @{ Row = 5;  Func = 'Get';            Fx = '=RedisUDFGet("{1}.key","{0}")' -f $h, $kp;                                                            Expected = 'hello_from_udf' },
        @{ Row = 6;  Func = 'Exists';         Fx = '=RedisUDFExists("{1}.key","{0}")' -f $h, $kp;                                                         Expected = '1' },
        @{ Row = 7;  Func = 'TTL';            Fx = '=RedisUDFTTL("{1}.key","{0}")' -f $h, $kp;                                                            Expected = '-1' },
        @{ Row = 8;  Func = 'SetJSON';        Fx = '=RedisUDFSetJSON("{1}.json",$F$2:$G$3,"{0}")' -f $h, $kp;                                              Expected = 'OK' },
        @{ Row = 9;  Func = 'JSONToMatrix';   Fx = '=INDEX(RedisUDFJSONToMatrix(RedisUDFGet("{1}.json","{0}"),""),2,1)' -f $h, $kp;                        Expected = '2' },
        @{ Row = 10; Func = 'HashSet';        Fx = '=RedisUDFHashSet("{1}.hash","campo1","valor1","{0}")' -f $h, $kp;                                      Expected = 'OK' },
        @{ Row = 11; Func = 'HashGet';        Fx = '=RedisUDFHashGet("{1}.hash","campo1","{0}")' -f $h, $kp;                                               Expected = 'valor1' },
        @{ Row = 12; Func = 'ChannelPublish'; Fx = '=RedisUDFChannelPublish("{1}.channel","ola_mundo","{0}")' -f $h, $kp;                                  Expected = '1 reader' },
        @{ Row = 13; Func = 'ChannelLatest';  Fx = '=RedisUDFChannelLatest("{1}.channel","{0}")' -f $h, $kp;                                               Expected = 'ola_mundo' },
        @{ Row = 14; Func = 'ConnectionCount';Fx = '=RedisUDFConnectionCount()';                                                                           Expected = $null },
        @{ Row = 15; Func = 'ExistsMultiples';Fx = '=INDEX(RedisUDFExistsMultiples({{"{1}.key","{1}.missing"}},"{0}"),1,2)' -f $h, $kp;                  Expected = '1' },
        @{ Row = 16; Func = 'TTLMultiples';   Fx = '=INDEX(RedisUDFTTLMultiples({{"{1}.key"}},"{0}"),1,2)' -f $h, $kp;                                  Expected = '-1' },
        @{ Row = 17; Func = 'HashGetFieldMultipleKeys'; Fx = '=INDEX(RedisUDFHashGetFieldMultipleKeys({{"{1}.hash","{1}.hash"}},"campo1","{0}"),2,2)' -f $h, $kp; Expected = 'valor1' },
        @{ Row = 18; Func = 'SetEx';          Fx = '=RedisUDFSetEx("{1}.tmp","x","100","{0}")' -f $h, $kp;                                                 Expected = 'OK' },
        @{ Row = 19; Func = 'TTL';            Fx = '=RedisUDFTTL("{1}.tmp","{0}")' -f $h, $kp;                                                             Expected = '100' },
        @{ Row = 20; Func = 'Del';            Fx = '=RedisUDFDel("{1}.neverexists","{0}")' -f $h, $kp;                                                     Expected = '0' },
        @{ Row = 21; Func = 'Incr';           Fx = '=RedisUDFIncr("{1}.counter","{0}")' -f $h, $kp;                                                        Expected = $null },
        @{ Row = 22; Func = 'ListPushRight';  Fx = '=RedisUDFListPushRight("{1}.list","a","{0}")' -f $h, $kp;                                              Expected = $null },
        @{ Row = 23; Func = 'ListRange';      Fx = '=INDEX(RedisUDFListRange("{1}.list",0,-1,"{0}"),1,1)' -f $h, $kp;                                      Expected = 'a' },
        @{ Row = 24; Func = 'SetAdd';         Fx = '=RedisUDFSetAdd("{1}.set","x","{0}")' -f $h, $kp;                                                      Expected = $null },
        @{ Row = 25; Func = 'SetMembers';     Fx = '=INDEX(RedisUDFSetMembers("{1}.set","{0}"),1,1)' -f $h, $kp;                                           Expected = 'x' },
        @{ Row = 26; Func = 'Set locale';     Fx = '=RedisUDFSet("{1}.locale",$H$2,"{0}")' -f $h, $kp;                                               Expected = 'OK' },
        @{ Row = 27; Func = 'Get locale';     Fx = '=RedisUDFGet("{1}.locale","{0}")' -f $h, $kp;                                                    Expected = '67000.5' },
        @{ Row = 28; Func = 'Set numeric key'; Fx = '=RedisUDFSet($H$2,"dec-key","{0}")' -f $h;                                                       Expected = 'OK' },
        @{ Row = 29; Func = 'Get numeric key'; Fx = '=RedisUDFGet($H$2,"{0}")' -f $h;                                                                 Expected = 'dec-key' },
        @{ Row = 30; Func = 'HashDel missing'; Fx = '=RedisUDFHashDel("{1}.hash","missing","{0}")' -f $h, $kp;                                             Expected = '0' },
        @{ Row = 31; Func = 'SetRemove missing'; Fx = '=RedisUDFSetRemove("{1}.set","missing","{0}")' -f $h, $kp;                                          Expected = '0' },
        @{ Row = 32; Func = 'Type missing'; Fx = '=RedisUDFType("{1}.missingkey","{0}")' -f $h, $kp;                                                       Expected = 'none' },
        @{ Row = 33; Func = 'Type string'; Fx = '=RedisUDFType("{1}.key","{0}")' -f $h, $kp;                                                               Expected = 'string' },
        @{ Row = 34; Func = 'Type list'; Fx = '=RedisUDFType("{1}.list","{0}")' -f $h, $kp;                                                                Expected = 'list' },
        @{ Row = 35; Func = 'ListPopRight empty'; Fx = '=IF(RedisUDFListPopRight("{1}.emptylist","{0}")="","empty","not empty")' -f $h, $kp;              Expected = 'empty' },
        @{ Row = 36; Func = 'ListPopLeft empty';  Fx = '=IF(RedisUDFListPopLeft("{1}.emptylist","{0}")="","empty","not empty")' -f $h, $kp;               Expected = 'empty' },
        @{ Row = 37; Func = 'Rename missing';     Fx = '=RedisUDFRename("{1}.missingrename","{1}.renamed","{0}")' -f $h, $kp;                               Expected = $null },
        @{ Row = 38; Func = 'Get range arg';  Fx = '=IF(ISNUMBER(SEARCH("Error",RedisUDFGet($F$2:$G$2))),"error","no error")';                               Expected = 'error' },
        @{ Row = 39; Func = 'GetMultiple no keys'; Fx = '=RedisUDFGetMultiple("",FALSE)';                                                                   Expected = $null },
        @{ Row = 40; Func = 'ExistsMultiples 2x2'; Fx = '=INDEX(RedisUDFExistsMultiples({{"{1}.key","{1}.missing1";"{1}.missing2","{1}.missing3"}},"{0}"),2,1)' -f $h, $kp; Expected = $null },
        @{ Row = 41; Func = 'Keys blank pattern'; Fx = '=RedisUDFKeys($J$3)';                                                                                 Expected = $null },
        @{ Row = 42; Func = 'GetMultiple 2x2'; Fx = '=INDEX(RedisUDFGetMultiple($F$2:$G$3,TRUE,"{0}"),2,1)' -f $h; Expected = $null },
        # Regression (v1.2.5): SetKV must reject ranges with different cell counts
        # instead of silently writing only part of the data.
        @{ Row = 43; Func = 'SetKV mismatched'; Fx = '=RedisUDFSetKV($F$2:$G$2,$F$3)'; Expected = $null },
        # Regression (v1.2.6): the second identical PublishIfChanged call must be
        # suppressed ("No change"), so the two calls in this formula differ.
        @{ Row = 44; Func = 'PublishIfChanged'; Fx = '=RedisUDFChannelPublishIfChanged("{1}.rtd2","same","{0}")' -f $h, $kp; Expected = 'No change' },
        # Regression (v1.2.6): blank channels are rejected with a clear Error cell.
        @{ Row = 45; Func = 'Unsubscribe blank channel'; Fx = '=RedisUDFChannelUnsubscribe("   ")'; Expected = $null },
        # Pair-range positive layout (v1.2.7): a 2x2 range is read row by row
        # (F2:G3 = 1/linha1;2/linha2 -> fields "1"="linha1" and "2"="linha2").
        @{ Row = 46; Func = 'HashSetMultiple pairs'; Fx = '=RedisUDFHashSetMultiple("{1}.pairs",$F$2:$G$3,"{0}")' -f $h, $kp; Expected = $null },
        @{ Row = 47; Func = 'HashGet pairs'; Fx = '=RedisUDFHashGet("{1}.pairs","1","{0}")' -f $h, $kp; Expected = 'linha1' }
    )
    foreach ($item in $udfItems) {
        Set-Cell $udf $item.Row 1 $item.Func
        Set-Formula $udf ("B{0}" -f $item.Row) $item.Fx
        Set-Cell $udf $item.Row 3 $item.Expected
    }

    # ------------------------------------------------------------ RTD sheet ----
    $rtd = Invoke-ExcelAction { $script:Workbook.Worksheets.Add([System.Reflection.Missing]::Value, $udf) }
    $rtd.Name = 'RTD'
    Set-Cell $rtd 1 1 ("RedisExcel RTD tests (host " + $RedisHost + ")")
    Set-Cell $rtd 3 1 'Function'
    Set-Cell $rtd 3 2 'Live result'
    Set-Cell $rtd 3 3 'Expected'

    $rtdItems = @(
        @{ Row = 4;  Func = 'GET';                 Fx = '=RTD("RedisRtd",,"GET","{1}.key","{0}")' -f $h, $kp;               Expected = 'hello_from_udf' },
        @{ Row = 5;  Func = 'HGET';                Fx = '=RTD("RedisRtd",,"HGET","{1}.hash","campo1","{0}")' -f $h, $kp;     Expected = 'valor1' },
        @{ Row = 6;  Func = 'HGETALL';             Fx = '=RTD("RedisRtd",,"HGETALL","{1}.hash","{0}")' -f $h, $kp;          Expected = '{"campo1":"valor1"}' },
        @{ Row = 7;  Func = 'SUB';                 Fx = '=RTD("RedisRtd",,"SUB","{1}.rtd","{0}")' -f $h, $kp;              Expected = $null },
        @{ Row = 8;  Func = 'PSUB';                Fx = '=RTD("RedisRtd",,"PSUB","{1}.rtd*","{0}")' -f $h, $kp;            Expected = $null },
        @{ Row = 9;  Func = 'ConnectionCount';     Fx = '=RedisRTDConnectionCount()';                                  Expected = $null },
        @{ Row = 10; Func = 'TopicCount';          Fx = '=RedisRTDTopicCount()';                                       Expected = $null },
        @{ Row = 11; Func = 'SubscriptionCount';   Fx = '=RedisRTDSubscriptionCount()';                                Expected = $null },
        @{ Row = 12; Func = 'ChannelCount';        Fx = '=RedisRTDChannelCount()';                                     Expected = $null },
        @{ Row = 13; Func = 'DefaultHost';         Fx = '=RedisRTDDefaultHost()';                                      Expected = $null },
        @{ Row = 14; Func = 'ExcelUpdateInterval'; Fx = '=RedisRTDExcelUpdateInterval()';                              Expected = $null },
        @{ Row = 15; Func = 'RedisUpdateInterval'; Fx = '=RedisRTDRedisUpdateInterval()';                              Expected = $null },
        @{ Row = 16; Func = 'RealTimeUpdates';     Fx = '=RedisRTDRealTimeUpdates()';                                  Expected = $null },
        # Regression (v1.2.3): a topic with a missing key must come back as
        # #ERROR without aborting the polling of the other topics of the host;
        # empty/whitespace keys are valid Redis names and must keep working.
        @{ Row = 20; Func = 'GET missing key';     Fx = '=RTD("RedisRtd",,"GET")';                                     Expected = $null },
        @{ Row = 21; Func = 'GET whitespace key';  Fx = '=RTD("RedisRtd",,"GET"," ","{0}")' -f $h;                     Expected = 'ws_key_value' },
        # Regression (v1.2.7): a missing hash is valid empty JSON, not a sentinel.
        @{ Row = 22; Func = 'HGETALL missing hash'; Fx = '=RTD("RedisRtd",,"HGETALL","{1}.missinghash","{0}")' -f $h, $kp; Expected = $null }
    )
    foreach ($item in $rtdItems) {
        Set-Cell $rtd $item.Row 1 $item.Func
        Set-Formula $rtd ("B{0}" -f $item.Row) $item.Fx
        Set-Cell $rtd $item.Row 3 $item.Expected
    }

    # Optional real production channels (read-only): SUB row 17, PSUB row 18.
    if ($RealChannel) {
        Set-Cell $rtd 17 1 ("SUB real (" + $RealChannel + ")")
        Set-Formula $rtd 'B17' ('=RTD("RedisRtd",,"SUB","{0}","{1}")' -f $RealChannel, $h)
    }
    if ($RealPattern) {
        Set-Cell $rtd 18 1 ("PSUB real (" + $RealPattern + ")")
        Set-Formula $rtd 'B18' ('=RTD("RedisRtd",,"PSUB","{0}","{1}")' -f $RealPattern, $h)
    }

    # ------------------------------------------------------- first asserts ----
    Invoke-ExcelAction { $script:Excel.CalculateFull() } | Out-Null
    Start-Sleep -Milliseconds 300
    Invoke-ExcelAction { $script:Excel.CalculateFull() } | Out-Null

    Check (Wait-CellText $udf 'B4' $script:WriteAck)                     'UDF Set returns the write ack'
    Check (Wait-CellText $udf 'B5' 'hello_from_udf')         'UDF Get returns the value'
    Check (Wait-CellText $udf 'B6' '1')                      'UDF Exists returns 1'
    Check (Wait-CellText $udf 'B7' '-1')                     'UDF TTL returns -1 (no expiry)'
    Check (Wait-CellText $udf 'B8' $script:WriteAck)                     'UDF SetJSON returns the write ack'
    Check (Wait-CellText $udf 'B9' '2')                      'UDF JSONToMatrix index [2,1] is 2'
    Check (Wait-CellText $udf 'B10' $script:WriteAck)                    'UDF HashSet returns the write ack'
    Check (Wait-CellText $udf 'B11' 'valor1')                'UDF HashGet returns the value'
    Check (Wait-CellText $udf 'B12' $script:ChannelPublishExpected)          'UDF ChannelPublish reports readers / the FF ack'
    Check (Wait-CellText $udf 'B13' 'ola_mundo')             'UDF ChannelLatest received the published message'
    Check (Wait-CellNumberMin $udf 'B14' 1)                  'UDF ConnectionCount >= 1'
    Check (Wait-CellText $udf 'B15' '1')                     'UDF ExistsMultiples (pipelined) first key exists'
    Check (Wait-CellText $udf 'B16' '-1')                    'UDF TTLMultiples (pipelined) returns -1'
    Check (Wait-CellText $udf 'B17' 'valor1')                'UDF HashGetFieldMultipleKeys (pipelined) returns the value'
    Check (Wait-CellText $udf 'B18' $script:WriteAck)                    'UDF SetEx returns the write ack'
    $ttlOk = $false
    $ttlLast = ''
    $ttlDeadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $ttlDeadline) {
        $ttlValue = 0.0
        $ttlLast = (Get-CellText $udf 'B19').Replace(',', '.')
        if ([double]::TryParse($ttlLast, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ttlValue) -and $ttlValue -gt 0 -and $ttlValue -le 100) { $ttlOk = $true; break }
        Start-Sleep -Milliseconds 250
    }
    if (-not $ttlOk) { Write-Host ("      B19 = '{0}'" -f $ttlLast) -ForegroundColor DarkGray }
    Check $ttlOk 'UDF TTL sees the SetEx expiry'
    Check (Wait-CellText $udf 'B20' $script:ZeroReplyExpected)                     'UDF Del reports 0 / the FF-all ack for a missing key'
    Check (Wait-CellRegex $udf 'B21' $script:IntReplyPattern)                'UDF Incr returns an integer / the FF-all ack'
    Check (Wait-CellRegex $udf 'B22' $script:ListPushReplyPattern)                'UDF ListPushRight returns an integer / the FF ack'
    Check (Wait-CellText $udf 'B23' 'a')                     'UDF ListRange returns the first element'
    Check (Wait-CellRegex $udf 'B24' $script:IntReplyPattern)                'UDF SetAdd returns an integer / the FF-all ack'
    Check (Wait-CellText $udf 'B25' 'x')                     'UDF SetMembers returns the member'
    Check (Wait-CellText $udf 'B26' $script:WriteAck)                        'UDF Set stores numeric cells invariantly'
    Check (Wait-CellText $udf 'B27' '67000.5')                   'UDF Get returns the invariant number'
    Check (Wait-CellText $udf 'B28' $script:WriteAck)                     'UDF Set stores numeric keys invariantly'
    Check (Wait-CellText $udf 'B29' 'dec-key')                'UDF Get reads numeric keys invariantly'
    Check (Wait-CellText $udf 'B30' $script:ZeroReplyExpected)                      'UDF HashDel reports 0 / the FF-all ack for a missing field'
    Check (Wait-CellText $udf 'B31' $script:ZeroReplyExpected)                      'UDF SetRemove reports 0 / the FF-all ack for a missing member'
    Check (Wait-CellText $udf 'B32' 'none')                   'UDF Type returns none for a missing key'
    Check (Wait-CellText $udf 'B33' 'string')                 'UDF Type returns string'
    Check (Wait-CellText $udf 'B34' 'list')                   'UDF Type returns list'
    Check (Wait-CellText $udf 'B35' $script:ListPopEmptyExpected)                  'UDF ListPopRight reports an empty pop / the FF-all ack'
    Check (Wait-CellText $udf 'B36' $script:ListPopEmptyExpected)                  'UDF ListPopLeft reports an empty pop / the FF-all ack'
    Check (Wait-CellRegex $udf 'B37' $script:RenameMissingPattern)                'UDF Rename reports the mode-appropriate result for a missing key'
    Check (Wait-CellText $udf 'B38' 'error')                  'UDF scalar argument rejects a multi-cell range'
    Check (Wait-CellRegex $udf 'B39' '^Error')                'UDF GetMultiple returns Error when no valid key remains'
    # B40 flattens the 2x2 input row-major into a 4x2 echo matrix: element (2,1)
    # is the SECOND key in row-major order, so an exact "missing1" match proves
    # the order (an Error cell would fail it too).
    Check (Wait-CellText $udf 'B40' "$KeyPrefix.missing1")    'UDF multi-key functions flatten a 2x2 range row-major'
    Check (Wait-CellRegex $udf 'B41' '^Error')                'UDF Keys rejects a blank pattern'
    Check (Wait-CellText $udf 'B42' 'linha1')                 'UDF GetMultiple flattens a 2x2 range row-major'
    Check (Wait-CellRegex $udf 'B43' '^Error')                'UDF SetKV rejects mismatched key/value counts'
    Check (Wait-CellRegex $udf 'B45' '^Error')                'UDF ChannelUnsubscribe rejects a blank channel'
    Check (Wait-CellText $udf 'B47' 'linha1')                 'UDF HashSetMultiple pairs a 2x2 range row-by-row'

    # ------------------------------- v1.3.0 NonVolatile twins (rows 48/49) ----
    # A ...NonVolatile write must run on entry only: a plain recalculation
    # (F9/edit) must not re-send it (Ctrl+Alt+F9 does, so no full calculation
    # below). The rows are entered AFTER the initial full calculations so the
    # counter starts at 1; the plain recalculations below must not touch it.
    Set-Cell $udf 48 1 'SetNonVolatile'
    Set-Cell $udf 49 1 'IncrNonVolatile'
    # Start from a clean counter so "1" below means "evaluated exactly once"
    # even when the test runs repeatedly against the same Redis.
    Invoke-RedisCli @('DEL', "$KeyPrefix.nvkey", "$KeyPrefix.nvcounter") | Out-Null
    Set-Formula $udf 'B48' ('=RedisUDFSetNonVolatile("{1}.nvkey","v1","{0}")' -f $h, $kp)
    Set-Formula $udf 'B49' ('=RedisUDFIncrNonVolatile("{1}.nvcounter","{0}")' -f $h, $kp)
    # Entry evaluation. Wait for both cells to settle (with AsyncWrites they
    # show Excel's pending marker until the queued write completed) before
    # touching Redis, so the CLI cannot race the entry writes.
    Invoke-ExcelAction { $udf.Calculate() } | Out-Null
    $nvSetOk = Wait-CellNotEmpty $udf 'B48' 20
    $nvIncrOk = Wait-CellNotEmpty $udf 'B49' 20
    $nvSetCell = Get-CellText $udf 'B48'
    $nvIncrCell = Get-CellText $udf 'B49'
    # Wait-CellNotEmpty accepts any non-# text, so an "Error: ..." value would
    # pass it: reject error/error-cell text explicitly, exactly like the
    # neighbouring check below.
    $nvEntryComputed = $nvSetOk -and $nvIncrOk `
        -and -not $nvSetCell.StartsWith('Error') -and -not $nvIncrCell.StartsWith('Error') `
        -and -not $nvSetCell.StartsWith('#') -and -not $nvIncrCell.StartsWith('#')
    Check $nvEntryComputed ("UDF NonVolatile twins evaluated on entry (B48='" + $nvSetCell + "', B49='" + $nvIncrCell + "')")
    Write-Host ("      DIAG nv-entry: B48='" + (Get-CellText $udf 'B48') + "' B49='" + (Get-CellText $udf 'B49') + "' nvkey='" + (Get-RedisValue @('GET', "$KeyPrefix.nvkey")) + "' nvcounter='" + (Get-RedisValue @('GET', "$KeyPrefix.nvcounter")) + "'") -ForegroundColor DarkGray

    # (a) A recalculation must not re-run SetNonVolatile: nvkey stays v2 and
    # the B48 cell keeps the value it computed on entry (no Error).
    Invoke-RedisCli @('SET', "$KeyPrefix.nvkey", 'v2') | Out-Null
    Invoke-ExcelAction { $udf.Calculate() } | Out-Null
    $nvKey = Get-RedisValue @('GET', "$KeyPrefix.nvkey")
    Check ($nvKey -eq 'v2') 'UDF SetNonVolatile does not re-run on a worksheet recalculation'
    $nvCellText = Get-CellText $udf 'B48'
    Check (-not [string]::IsNullOrWhiteSpace($nvCellText) -and -not $nvCellText.StartsWith('#') -and -not $nvCellText.StartsWith('Error')) 'UDF SetNonVolatile cell was computed (not an error)'

    # (b) Two more plain recalculations must not increment the counter again
    # (a volatile Incr would have reached 4 by now: entry + three recalcs).
    Invoke-ExcelAction { $udf.Calculate() } | Out-Null
    Invoke-ExcelAction { $udf.Calculate() } | Out-Null
    $nvCounter = Get-RedisValue @('GET', "$KeyPrefix.nvcounter")
    $nvKeyAfter = Get-RedisValue @('GET', "$KeyPrefix.nvkey")
    Check ($nvCounter -eq '1') ("UDF IncrNonVolatile evaluated once only (recalcs do not increment; nvcounter='" + $nvCounter + "', nvkey='" + $nvKeyAfter + "')")

    # ------------------------------ v1.4.0 async-mode checks (rows 50-53) ----
    # Gated on -AsyncWrites (audit-4 A/B/C): (A) the pending marker and exactly
    # one delivery per registered call, (B) the calling cell is part of the
    # async identity (two identical formulas must both write), (C) an argument
    # change dispatches a new write (last value wins).
    if ($AsyncWrites) {
        # Manual calculation keeps the pending state observable and makes the
        # delivery recalculation explicit; the previous mode is restored in the
        # finally even when a check fails. Check C runs in normal (automatic)
        # calculation mode.
        $previousCalculation = Invoke-ExcelAction { $script:Excel.Calculation }
        try {
            Invoke-ExcelAction { $script:Excel.Calculation = -4135 } | Out-Null

            # The rows below must stay free: earlier sections fill the sheet
            # top-down, so a future check reaching this block would silently
            # overwrite these labels/formulas.
            $asyncBlockEmpty = $true
            foreach ($row in 50..53) {
                foreach ($col in 'A', 'B', 'C', 'D', 'E', 'F') {
                    if (-not [string]::IsNullOrWhiteSpace((Get-CellText $udf "$col$row"))) { $asyncBlockEmpty = $false }
                }
            }
            Check $asyncBlockEmpty 'async-mode check block (rows 50-53) is unused'

            # (A) Pending marker + single delivery (row 50).
            Invoke-RedisCli @('DEL', "$KeyPrefix.asyncmarker") | Out-Null
            Set-Cell $udf 50 1 'Incr (async pending/single delivery)'
            Set-Formula $udf 'B50' ('=RedisUDFIncr("{1}.asyncmarker","{0}")' -f $h, $kp)
            # The pending result is Excel's N/A error, detected through the
            # object model (WorksheetFunction.IsNA) so the check is
            # locale-independent (#N/A, #N/D, #NV, ...).
            $pendingOk = Wait-CellIsNA $udf 'B50' 5
            Check $pendingOk ("Async pending marker shown for the queued Incr (B50='" + (Get-CellText $udf 'B50') + "')")
            # The queued write must reach Redis while the cell still shows the
            # pending marker (manual mode suppresses the delivery recalculation).
            $markerOk = Wait-RedisValue @('GET', "$KeyPrefix.asyncmarker") '1' 5
            $markerCell = Get-CellText $udf 'B50'
            $markerStillNa = Test-CellIsNA $udf 'B50'
            $markerRedis = Get-RedisValue @('GET', "$KeyPrefix.asyncmarker")
            Check ($markerOk -and $markerStillNa) ("Async Incr reached Redis while the cell was still pending (redis='" + $markerRedis + "', cell='" + $markerCell + "', isNA=" + $markerStillNa + ")")
            # Force the delivery recalculation: it must return the cached result
            # (the mode-appropriate reply, not a fresh write); the single retry
            # only covers the completion notification racing this Calculate.
            $asyncIncrReply = if ($script:IsFireForgetAll) { 'OK (fire and forget: all)' } else { '1' }
            Invoke-ExcelAction { $udf.Range('B50').Calculate() } | Out-Null
            $settledOk = Wait-CellText $udf 'B50' $asyncIncrReply 10
            if (-not $settledOk) {
                Invoke-ExcelAction { $udf.Range('B50').Calculate() } | Out-Null
                $settledOk = Wait-CellText $udf 'B50' $asyncIncrReply 10
            }
            $markerAfter = Get-RedisValue @('GET', "$KeyPrefix.asyncmarker")
            Check ($settledOk -and $markerAfter -eq '1') ("Async delivery recalculation settles the cell to '" + $asyncIncrReply + "' without re-running the write (cell='" + (Get-CellText $udf 'B50') + "', redis='" + $markerAfter + "')")

            # (B) Two cells, IDENTICAL formula, both must write (rows 51/52):
            # the calling cell is part of the async identity, so a collapsed
            # identity would leave only one ListPushRight (LLEN=1).
            Invoke-RedisCli @('DEL', "$KeyPrefix.asyncdup") | Out-Null
            Set-Cell $udf 51 1 'ListPushRight A (async identity)'
            Set-Cell $udf 52 1 'ListPushRight B (async identity)'
            $dupFormula = '=RedisUDFListPushRight("{1}.asyncdup","x","{0}")' -f $h, $kp
            Set-Formula $udf 'B51' $dupFormula
            Set-Formula $udf 'B52' $dupFormula
            $dupOk = Wait-RedisValue @('LLEN', "$KeyPrefix.asyncdup") '2' 10
            # Re-read twice after settling: a late duplicate write would drift
            # the final length from 2 (two equal reads one second apart).
            Start-Sleep -Milliseconds 300
            $dupLen = Get-RedisValue @('LLEN', "$KeyPrefix.asyncdup")
            $dupLen2 = $dupLen
            if ($dupLen -eq '2') {
                Start-Sleep -Milliseconds 1000
                $dupLen2 = Get-RedisValue @('LLEN', "$KeyPrefix.asyncdup")
            }
            $dupOk = $dupOk -and ($dupLen -eq '2') -and ($dupLen2 -eq '2')
            if (-not $dupOk) {
                Write-Host ("      DIAG asyncdup: B51='" + (Get-CellText $udf 'B51') + "' B52='" + (Get-CellText $udf 'B52') + "'") -ForegroundColor DarkGray
            }
            Check $dupOk ("Two cells with the identical async formula both wrote (LLEN='" + $dupLen + "/" + $dupLen2 + "')")

            # Negative half of the identity contract: the SAME cell re-evaluated
            # must not dispatch a second write. One cell, one push (LLEN=1); a
            # forced recalculation returns the cached result, so it stays 1.
            Invoke-RedisCli @('DEL', "$KeyPrefix.asyncsingle") | Out-Null
            Set-Formula $udf 'B51' ('=RedisUDFListPushRight("{1}.asyncsingle","y","{0}")' -f $h, $kp)
            $singleOk = Wait-RedisValue @('LLEN', "$KeyPrefix.asyncsingle") '1' 10
            Invoke-ExcelAction { $udf.Range('B51').Calculate() } | Out-Null
            Start-Sleep -Milliseconds 400
            $singleLen = Get-RedisValue @('LLEN', "$KeyPrefix.asyncsingle")
            Check ($singleOk -and $singleLen -eq '1') ("A single cell re-evaluated keeps exactly one async write (LLEN='" + $singleLen + "')")

            # (C) Argument change dispatches a new write (row 53), back in
            # normal (automatic) calculation mode: the F53 edit must
            # recalculate B53 and dispatch with the new argument value.
            Invoke-ExcelAction { $script:Excel.Calculation = -4105 } | Out-Null
            Invoke-RedisCli @('DEL', "$KeyPrefix.asyncarg") | Out-Null
            Set-Cell $udf 53 1 'Set (async argument change)'
            Set-Cell $udf 53 6 'v1'
            Set-Formula $udf 'B53' ('=RedisUDFSet("{1}.asyncarg",$F$53,"{0}")' -f $h, $kp)
            $argV1Ok = Wait-RedisValue @('GET', "$KeyPrefix.asyncarg") 'v1' 10
            Set-Cell $udf 53 6 'v2'
            $argV2Ok = Wait-RedisValue @('GET', "$KeyPrefix.asyncarg") 'v2' 10
            $argFinal = Get-RedisValue @('GET', "$KeyPrefix.asyncarg")
            Check ($argV1Ok -and $argV2Ok -and $argFinal -eq 'v2') ("Async argument change dispatches a new write and the last value wins (v1=" + $argV1Ok + ", v2=" + $argV2Ok + ", final='" + $argFinal + "')")
        }
        finally {
            try {
                Invoke-ExcelAction { $script:Excel.Calculation = $previousCalculation } | Out-Null
            }
            catch {
                Write-Host ("WARNING: could not restore Application.Calculation: " + $_.Exception.Message) -ForegroundColor Red
                $script:Failures++
            }
        }
    }

    Check (Wait-CellText $rtd 'B4' 'hello_from_udf')         'RTD GET returns the value'
    Check (Wait-CellText $rtd 'B5' 'valor1')                 'RTD HGET returns the value'
    Check (Wait-CellText $rtd 'B6' '{"campo1":"valor1"}')    'RTD HGETALL returns valid JSON'
    Check (Wait-CellText $rtd 'B22' '{}')                    'RTD HGETALL returns {} for a missing hash'

    Invoke-RedisCli @('SET', ' ', 'ws_key_value') | Out-Null
    Invoke-RedisCli @('PUBLISH', "$KeyPrefix.rtd", 'ALTA') | Out-Null
    Check (Wait-CellText $rtd 'B7' 'ALTA')                   'RTD SUB received the published message'
    Check (Wait-CellText $rtd 'B8' 'ALTA')                   'RTD PSUB received the published message'
    Check (Wait-CellNumberMin $rtd 'B9' 1)                   'RTD ConnectionCount >= 1'
    Check (Wait-CellNumberMin $rtd 'B10' 5)                  'RTD TopicCount >= 5'
    Check (Wait-CellNumberMin $rtd 'B11' 2)                  'RTD SubscriptionCount >= 2'
    Check (Wait-CellNumberMin $rtd 'B12' 1)                  'RTD ChannelCount >= 1'
    # Rows 13-16 were built but never asserted (coverage gap): the status helpers
    # must report the configured default host and positive intervals, and the
    # real-time flag must render as a boolean (TRUE/FALSE, locale-tolerant).
    Check (Wait-CellNotEmpty $rtd 'B13' 10)                  'RTD DefaultHost reports a non-empty host'
    # B13 is a status helper: it must return the exact configured default host
    # text (not a substring match, which would also accept an error cell that
    # merely contains the host) and must not be an Excel error cell.
    $defaultHostText = Get-CellText $rtd 'B13'
    $expectedDefaultHost = 'localhost:6379,password=,defaultDatabase=0,ssl=False,abortConnect=False'
    Check ($defaultHostText -eq $expectedDefaultHost) ("RTD DefaultHost matches the configured host (B13='" + $defaultHostText + "')")
    Check (-not $defaultHostText.StartsWith('#') -and -not $defaultHostText.StartsWith('Error')) ("RTD DefaultHost is a computed value, not an error cell (B13='" + $defaultHostText + "')")
    Check (Wait-CellNumberMin $rtd 'B14' 1)                  'RTD ExcelUpdateInterval >= 1'
    Check (Wait-CellNumberMin $rtd 'B15' 1)                  'RTD RedisUpdateInterval >= 1'
    Check (Wait-CellRegex $rtd 'B16' '(?i)^(true|false|verdadeiro|falso)$') 'RTD RealTimeUpdates reports a boolean'
    Check (Wait-CellText $rtd 'B21' 'ws_key_value')          'RTD GET accepts a whitespace-only key'
    Check (Wait-CellRegex $rtd 'B20' '^#ERROR')              'RTD GET without a key returns #ERROR without freezing the host'

    # PublishIfChanged only marks delivered payloads; the PSUB topic above is a
    # live reader, so two forced recalculations settle the cell on "No change"
    # (with zero readers every recalculation retries the publish).
    Invoke-ExcelAction { $udf.Calculate() } | Out-Null
    Start-Sleep -Milliseconds 400
    Invoke-ExcelAction { $udf.Calculate() } | Out-Null
    Check (Wait-CellText $udf 'B44' 'No change')             'UDF PublishIfChanged suppresses an unchanged payload'
    Invoke-RedisCli @('DEL', ' ') | Out-Null

    # A polled GET topic must pick up a value changed after registration. The
    # volatile UDF sheet rewrites {prefix}.key on every recalculation (row 4
    # Set), so a recalculation triggered around the first push can revert the
    # key ~1s after it was written; re-issue the SET until the RTD cell shows
    # the new value (same pattern as Publish-Until-Cell).
    $seen = $false
    $readBack = ''
    $b4Deadline = (Get-Date).AddSeconds(30)
    while (-not $seen -and (Get-Date) -lt $b4Deadline) {
        Invoke-RedisCli @('SET', "$KeyPrefix.key", 'hello_v2') | Out-Null
        $readBack = Get-RedisValue @('GET', "$KeyPrefix.key")
        $seen = Wait-CellText $rtd 'B4' 'hello_v2' 3
    }
    Check $seen ("RTD GET picks up a changed value (last read back '" + $readBack + "')")
    Invoke-RedisCli @('SET', "$KeyPrefix.key", 'hello_from_udf') | Out-Null

    if ($RealChannel) {
        Check (Wait-CellNotEmpty $rtd 'B17' 30)              'RTD SUB received live data from the real channel'
    }
    if ($RealPattern) {
        Check (Wait-CellNotEmpty $rtd 'B18' 30)              'RTD PSUB received live data from the real pattern'
    }

    # ------------------------------------------------------------- save it ----
    # Save to a temp file first and copy it into place afterwards, so a failed
    # SaveAs can never delete the committed sample workbook.
    $tempOut = Join-Path $env:TEMP 'RedisExcel.Test.xlsx'
    $outPath = $tempOut
    if ($isLocalHost -and -not $AsyncWrites -and -not $RealChannel -and -not $RealPattern) {
        $outDir = Join-Path $RepoRoot 'test'
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
        $outPath = Join-Path $outDir 'RedisExcel.Test.xlsx'
    }
    else {
        # Remote hosts never write into the repository, an async run adds
        # scratch cells (rows 50-53) that do not belong in the committed
        # sample, and -RealChannel/-RealPattern put real data into the RTD
        # sheet's B17/B18 cells, which the metadata sanitizer does not touch:
        # keep all of those runs in %TEMP%.
        Write-Host "The workbook will not be saved into the repository (remote host, -AsyncWrites or -RealChannel/-RealPattern)." -ForegroundColor DarkGray
    }
    Remove-Item -LiteralPath $tempOut -Force -ErrorAction SilentlyContinue
    Invoke-ExcelAction { $script:Workbook.SaveAs($tempOut, 51) } | Out-Null
    if ($tempOut -ne $outPath) {
        # Stage next to the target and replace with Move-Item, so the committed
        # sample is never left half-written if this process dies mid-copy.
        $stagedPath = "$outPath.tmp"
        Copy-Item -Path $tempOut -Destination $stagedPath -Force
        # Strip the local save path and personal metadata from the staged copy
        # BEFORE it replaces the committed sample, so a failed sanitize leaves
        # the previous sample untouched. Remote runs keep the workbook in
        # %TEMP% (still open in Excel and deleted in the cleanup), so there is
        # nothing to sanitize there.
        Remove-WorkbookMetadata -Path $stagedPath
        Move-Item -Path $stagedPath -Destination $outPath -Force
    }
    Check (Test-Path $outPath) ("test workbook saved to " + $outPath)

    # ------------------------------------ regression: copy of the workbook ----
    # The reported bug: copying/closing a workbook silently killed the Pub/Sub
    # subscriptions of the other workbook using the same channel.
    $copyPath = Join-Path $env:TEMP 'RedisExcel.Test.Copy.xlsx'
    Remove-Item -LiteralPath $copyPath -Force -ErrorAction SilentlyContinue
    Invoke-ExcelAction { $script:Workbook.SaveCopyAs($copyPath) } | Out-Null
    # Workbooks.Open can return null transiently right after SaveCopyAs (the
    # file may still be scanned/locked); retry before giving up and report the
    # file state when it never opens.
    $copy = $null
    $copyDeadline = (Get-Date).AddSeconds(30)
    while (-not $copy -and (Get-Date) -lt $copyDeadline) {
        $copy = Invoke-ExcelAction { $script:Excel.Workbooks.Open($copyPath) }
        if (-not $copy) {
            Write-Host '      DIAG copy open returned null; retrying...' -ForegroundColor DarkGray
            Start-Sleep -Milliseconds 500
        }
    }
    if (-not $copy) {
        $copyInfo = Get-Item $copyPath -ErrorAction SilentlyContinue
        throw ("copy workbook did not open: exists=" + (Test-Path $copyPath) + " size=" + $(if ($copyInfo) { $copyInfo.Length } else { 'n/a' }))
    }
    # The workbook object can come back before its object model answers, and a
    # busy Excel can reject the call (RPC_E_CALL_REJECTED); keep asking until
    # the RTD sheet resolves instead of aborting the whole run.
    $copyRtd = $null
    $copyRtdTries = 0
    $copyRtdDeadline = (Get-Date).AddSeconds(30)
    while (-not $copyRtd -and (Get-Date) -lt $copyRtdDeadline) {
        $copyRtdTries++
        try { $copyRtd = Invoke-ExcelAction { $copy.Worksheets.Item('RTD') } }
        catch { $copyRtd = $null }
        if (-not $copyRtd) {
            if ($copyRtdTries -eq 1) {
                Write-Host '      DIAG copy RTD sheet not ready; retrying...' -ForegroundColor DarkGray
            }
            Start-Sleep -Milliseconds 300
        }
    }
    if (-not $copyRtd) {
        throw "copy workbook has no usable RTD sheet (Worksheets.Item('RTD') never answered)"
    }

    Check (Publish-Until-Cell "$KeyPrefix.rtd" 'copy-1' $rtd 'B7' 'copy-1')        'original received copy-1'
    Check (Publish-Until-Cell "$KeyPrefix.rtd" 'copy-1' $copyRtd 'B7' 'copy-1')    'copy received copy-1'

    # Duplicate the RTD sheet inside the copy: new topics for the same channel.
    Invoke-ExcelAction { $copyRtd.Copy($copy.Worksheets.Item($copy.Worksheets.Count)) } | Out-Null
    # Pick the duplicated sheet: prefer the sheet copied just now (name starts
    # with 'RTD ('), rejecting the original 'RTD' sheet; only then fall back.
    # The whole selection runs inside Invoke-ExcelAction so transient COM
    # rejections are retried.
    $dup = Invoke-ExcelAction {
        $sheet = $script:Excel.ActiveSheet
        if ($sheet -and $sheet.Name -notlike 'RTD (*') {
            $sheet = @($copy.Worksheets | Where-Object { $_.Name -like 'RTD (*' })[0]
        }
        if (-not $sheet -or $sheet.Name -eq 'RTD') {
            # Fallback: only accept the second sheet when it is not the original
            # 'RTD' sheet; otherwise the 'duplicated sheet receives' check would
            # silently test the original sheet and pass without testing the
            # duplicated one.
            $candidate = $null
            try { $candidate = $copy.Worksheets.Item(2) } catch { $candidate = $null }
            if ($candidate -and $candidate.Name -ne 'RTD') { $sheet = $candidate }
        }
        if (-not $sheet -or $sheet.Name -eq 'RTD') {
            $names = (@($copy.Worksheets | ForEach-Object { $_.Name }) -join ', ')
            throw ("no duplicated RTD sheet found (need a sheet whose name is not 'RTD'); workbook sheets: " + $names)
        }
        $sheet
    }
    Check (Publish-Until-Cell "$KeyPrefix.rtd" 'copy-2' $rtd 'B7' 'copy-2')        'original keeps receiving after a sheet is copied'
    Check (Publish-Until-Cell "$KeyPrefix.rtd" 'copy-2' $copyRtd 'B7' 'copy-2')    'duplicated workbook keeps receiving'
    Check (Publish-Until-Cell "$KeyPrefix.rtd" 'copy-2' $dup 'B7' 'copy-2')        'duplicated sheet receives'

    Invoke-ExcelAction { $copy.Close($false) } | Out-Null
    Start-Sleep -Milliseconds 500
    Check (Publish-Until-Cell "$KeyPrefix.rtd" 'copy-3' $rtd 'B7' 'copy-3')        'original keeps receiving after the copy workbook is closed'

    # ---------------------------------------- regression: connection blip ----
    if ($allowClientKill) {
        Invoke-RedisCli @('CLIENT', 'KILL', 'TYPE', 'pubsub') | Out-Null
        Start-Sleep -Seconds 1
        Check (Publish-Until-Cell "$KeyPrefix.rtd" 'reconnect-1' $rtd 'B7' 'reconnect-1' 30)  'subscriptions recover after the Pub/Sub connection is killed'
    }
    else {
        Write-Host "SKIP  CLIENT KILL (not a local host, -SkipClientKill, or a custom -RedisCli); subscriptions not tested against a blip." -ForegroundColor DarkGray
    }

    # --------------------------- regression: Excel dies mid-stream ----
    # Kill ONLY this run's own hidden Excel: the victim is the EXCEL.EXE PID
    # outside the pre-run baseline captured before this script created its
    # instance, so a user window is never a candidate. While the RTD SUB/PSUB
    # and the UDF ChannelLatest listeners are live, then wait for the server to
    # drop every subscription. This proves a crashed add-in leaves no dangling
    # Pub/Sub state behind and a fresh client would start clean. Skipped with
    # -KeepExcelOpen, which intentionally leaves Excel running.
    if (-not $KeepExcelOpen) {
        $rtdChannel = "$KeyPrefix.rtd"       # RTD SUB topic (RTD sheet row 7)
        $udfChannel = "$KeyPrefix.channel"   # UDF ChannelLatest (UDF sheet row 13)

        # 1) The subscription state is live before the kill.
        $preRtd = Get-RedisPubsubCount @('PUBSUB', 'NUMSUB', $rtdChannel)
        $prePat = Get-RedisPubsubCount @('PUBSUB', 'NUMPAT')
        Write-Host ("      pre-kill : PUBSUB NUMSUB {0} = {1}; PUBSUB NUMPAT = {2} (PSUB {0}* live)" -f $rtdChannel, $preRtd, $prePat) -ForegroundColor DarkGray
        Check ($preRtd -ge 1) ('RTD subscription exists before the kill (NUMSUB ' + $rtdChannel + ' = ' + $preRtd + ', expected >= 1)')
        Check ($prePat -ge 1) ('PSUB pattern state exists before the kill (NUMPAT = ' + $prePat + ', expected >= 1)')

        # 2) Kill this run's hidden Excel. The baseline diff leaves exactly one
        #    candidate by design; any other count means we cannot tell which
        #    process is ours, so fail loudly and do not kill anything.
        $newExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $script:ExcelPidsBefore -notcontains $_.Id })
        if ($newExcel.Count -eq 1) {
            $killedPid = $newExcel[0].Id
            Write-Host ("      Stopping this run's hidden Excel (PID {0}) with Stop-Process -Force." -f $killedPid) -ForegroundColor DarkGray
            $script:ExcelKilled = $true
            try {
                Stop-Process -Id $killedPid -Force -ErrorAction Stop
            }
            catch {
                # The kill did not happen: restore the normal COM teardown.
                $script:ExcelKilled = $false
                Check $false ('could not kill the hidden Excel (PID ' + $killedPid + '): ' + $_.Exception.Message)
            }
        }
        else {
            Check $false ('expected exactly one EXCEL.EXE outside the pre-run baseline to kill, found ' + $newExcel.Count + ' (not killing anything)')
        }

        # 3) The killed process must tear its subscriptions down server-side
        #    (no Quit, no unsubscribe): the RTD SUB, the UDF channel listener
        #    and the RTD PSUB pattern all drop to zero, bounded at 20s.
        $postRtd = -1
        $postChan = -1
        $postPat = -1
        $cleanDeadline = (Get-Date).AddSeconds(20)
        while ((Get-Date) -lt $cleanDeadline) {
            $postRtd = Get-RedisPubsubCount @('PUBSUB', 'NUMSUB', $rtdChannel)
            $postChan = Get-RedisPubsubCount @('PUBSUB', 'NUMSUB', $udfChannel)
            $postPat = Get-RedisPubsubCount @('PUBSUB', 'NUMPAT')
            if ($postRtd -eq 0 -and $postChan -eq 0 -and $postPat -eq 0) { break }
            Start-Sleep -Milliseconds 250
        }
        Write-Host ("      post-kill: PUBSUB NUMSUB {0} = {1}; PUBSUB NUMSUB {2} = {3}; PUBSUB NUMPAT = {4}" -f $rtdChannel, $postRtd, $udfChannel, $postChan, $postPat) -ForegroundColor DarkGray
        Check ($postRtd -eq 0) ('no dangling RTD subscription after the crash (last NUMSUB ' + $rtdChannel + ' = ' + $postRtd + ', expected 0)')
        Check ($postChan -eq 0) ('no dangling UDF channel subscription after the crash (last NUMSUB ' + $udfChannel + ' = ' + $postChan + ', expected 0)')
        Check ($postPat -eq 0) ('no dangling pattern subscription after the crash (last NUMPAT = ' + $postPat + ', expected 0)')
    }
}
finally {
    if ($KeepExcelOpen) {
        if ($script:Excel -ne $null) {
            try { $script:Excel.Visible = $true } catch { }
            Write-Host 'Excel was left open (-KeepExcelOpen).' -ForegroundColor DarkYellow
        }
    }
    else {
        # The "Excel dies mid-stream" scenario killed this run's hidden Excel
        # with Stop-Process: the COM proxy is dead, so Workbook.Close/Quit/
        # ReleaseComObject and the wait-for-exit loop are skipped (COM calls
        # against the dead instance would only burn the retry budget). The
        # config restore and the temp-workbook cleanup below still run.
        if (-not $script:ExcelKilled) {
            try { if ($copy) { Invoke-ExcelAction { $copy.Close($false) } | Out-Null } } catch { }
            try { if ($script:Workbook) { Invoke-ExcelAction { $script:Workbook.Close($false) } | Out-Null } } catch { }
            if ($script:Excel -ne $null) {
                try { Invoke-ExcelAction { $script:Excel.Quit() } | Out-Null } catch { }
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Excel) | Out-Null
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            }

            # Wait for this run's hidden Excel to actually exit before returning: a
            # lingering instance (plus antivirus scanning its teardown) makes a
            # back-to-back run slow enough to trip the COM retries. Never kill
            # EXCEL.EXE blindly; wait bounded and report leftovers.
            $pending = @()
            $exitDeadline = (Get-Date).AddSeconds(30)
            while ((Get-Date) -lt $exitDeadline) {
                $pending = @(Get-Process EXCEL -ErrorAction SilentlyContinue |
                    Where-Object { $script:ExcelPidsBefore -notcontains $_.Id })
                if ($pending.Count -eq 0) { break }
                Start-Sleep -Milliseconds 250
            }
            if ($pending.Count -gt 0) {
                Write-Host ("WARNING: a hidden Excel instance did not exit within 30s (PIDs: " + (($pending | ForEach-Object { $_.Id }) -join ',') + "); back-to-back runs may be slow.") -ForegroundColor Yellow
            }
        }
        # Short cooldown so security software can finish scanning the files
        # this run created/removed before the next run starts.
        Start-Sleep -Seconds 2
    }
    # Restore the user's own RedisExcel.json (or remove the one written for this
    # run). The running Excel already read the config once at XLL load.
    try {
        if ($configBackup) {
            Copy-Item -LiteralPath $configBackup -Destination $configPath -Force
            Remove-Item -LiteralPath $configBackup -Force -ErrorAction SilentlyContinue
        }
        else {
            Remove-Item -LiteralPath $configPath -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host ("WARNING: could not restore {0}: {1}" -f $configPath, $_.Exception.Message) -ForegroundColor Red
    }

    # Clean up the temporary workbooks (copy + save-as staging) only when Excel
    # is gone; with -KeepExcelOpen the open workbooks still lock those files.
    # The committed sample in test\ is never removed here.
    $tempCopyPath = Join-Path $env:TEMP 'RedisExcel.Test.Copy.xlsx'
    $tempSavePath = Join-Path $env:TEMP 'RedisExcel.Test.xlsx'
    if ($KeepExcelOpen) {
        Write-Host ("Temporary workbooks kept: {0} ; {1}" -f $tempSavePath, $tempCopyPath) -ForegroundColor DarkYellow
    }
    else {
        Remove-Item -LiteralPath $tempSavePath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempCopyPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath (Join-Path (Join-Path $RepoRoot 'test') 'RedisExcel.Test.xlsx.tmp') -Force -ErrorAction SilentlyContinue
    }
}

if ($script:Failures -eq 0) {
    Write-Host 'ALL PASS' -ForegroundColor Green
    exit 0
}
Write-Host ("{0} FAILURE(S)" -f $script:Failures) -ForegroundColor Red
exit 1
