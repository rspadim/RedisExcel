<#
.SYNOPSIS
End-to-end test for the RedisExcel add-in using a real (hidden) Excel instance.

.DESCRIPTION
Builds the test workbook and asserts UDF/RTD results against a running Redis
server, including the v1.1.0 Pub/Sub regression scenario:

  1. Loads the packed XLL into Excel (RegisterXLL).
  2. Creates the UDF/RTD test sheets and saves test\RedisExcel.Test.xlsx
     (only for local hosts; remote hosts are saved to %TEMP% so the host is
     never committed to the repository).
  3. Asserts values (SET/GET/EXISTS/TTL/JSON/HASH, RTD GET/HGET/HGETALL/SUB/PSUB).
  4. Opens a COPY of the workbook, publishes messages, closes the copy and
     verifies the original workbook keeps receiving (the reported bug).
  5. Kills the Pub/Sub connections server-side (CLIENT KILL TYPE pubsub) and
     verifies the subscriptions recover automatically. This step only runs
     against local hosts; use -SkipClientKill to force it off.

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
    [switch]$KeepExcelOpen
)

$ErrorActionPreference = 'Stop'

$script:Failures = 0
$script:RedisExe = $null
$script:RedisPrefix = @()
$script:RedisArgs = @()
$script:Excel = $null
$script:Workbook = $null

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
    $container = ((& docker ps --filter name=redisexcel-test --format "{{.Names}}" 2>$null) -join '').Trim()
    if ($container -eq 'redisexcel-test') {
        $script:RedisExe = 'docker'
        $script:RedisPrefix = @('exec', 'redisexcel-test', 'redis-cli')
        return
    }
    throw "redis-cli not found and container 'redisexcel-test' is not running. Pass -RedisCli (e.g. 'docker exec my-redis redis-cli')."
}

function Test-ComBusyError($ErrorRecord) {
    $hr = $ErrorRecord.Exception.HResult
    return ($hr -eq -2147418111 -or $hr -eq -2147417846)
}

# Some security products intermittently break Excel COM property sets with a
# bogus InvalidCastException (e.g. "cannot convert Int32 to String"). Those
# failures are transient; retry them like a busy Excel.
function Test-RetryableError($ErrorRecord) {
    if ($ErrorRecord.Exception -is [System.InvalidCastException]) { return $true }
    return (Test-ComBusyError $ErrorRecord)
}

function Invoke-ExcelAction([scriptblock]$Action, [int]$Retries = 40) {
    for ($attempt = 0; $attempt -lt $Retries; $attempt++) {
        try { return (& $Action) }
        catch {
            if (-not (Test-RetryableError $_)) { throw }
            Start-Sleep -Milliseconds ([Math]::Min(150 * ($attempt + 1), 3000))
        }
    }
    return (& $Action)
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
    for ($attempt = 0; ; $attempt++) {
        try {
            $Sheet.Range($address).Value2 = $text
            if ($attempt -gt 0) {
                Write-Host ("      Set-Cell[{0}] succeeded after {1} retries (transient COM failure)" -f $address, $attempt) -ForegroundColor DarkYellow
            }
            return
        }
        catch {
            if ($attempt -ge 40 -or -not (Test-RetryableError $_)) {
                Write-Host ("Set-Cell[{0}] failed after {1} attempts: {2}" -f $address, $attempt, $_.Exception.Message) -ForegroundColor Red
                throw
            }
            Start-Sleep -Milliseconds ([Math]::Min(150 * ($attempt + 1), 3000))
        }
    }
}

function Set-Formula($Sheet, [string]$Address, [string]$Formula) {
    for ($attempt = 0; ; $attempt++) {
        try { $Sheet.Range($Address).Formula = $Formula; return }
        catch {
            if ($attempt -ge 40 -or -not (Test-RetryableError $_)) {
                Write-Host ("Set-Formula[{0}] failed after {1} attempts: {2}" -f $Address, $attempt, $_.Exception.Message) -ForegroundColor Red
                throw
            }
            Start-Sleep -Milliseconds ([Math]::Min(150 * ($attempt + 1), 3000))
        }
    }
}

function Get-CellText($Sheet, [string]$Address) {
    for ($attempt = 0; ; $attempt++) {
        try { return [string]$Sheet.Range($Address).Text }
        catch {
            if ($attempt -ge 40 -or -not (Test-RetryableError $_)) { throw }
            Start-Sleep -Milliseconds ([Math]::Min(150 * ($attempt + 1), 3000))
        }
    }
}

function Get-CellNumber($Sheet, [string]$Address) {
    try { return [double](Get-CellText $Sheet $Address) }
    catch { return [double]::NaN }
}

function Wait-CellText($Sheet, [string]$Address, [string]$Expected, [int]$TimeoutSeconds = 20) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $text = ''
    while ((Get-Date) -lt $deadline) {
        $text = Get-CellText $Sheet $Address
        if ($text -eq $Expected) { return $true }
        Start-Sleep -Milliseconds 300
    }
    Write-Host ("      {0} = '{1}' (expected '{2}')" -f $Address, $text, $Expected) -ForegroundColor DarkGray
    return $false
}

function Wait-CellNotEmpty($Sheet, [string]$Address, [int]$TimeoutSeconds = 20) {
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $text = ''
    while ((Get-Date) -lt $deadline) {
        $text = Get-CellText $Sheet $Address
        if (-not [string]::IsNullOrWhiteSpace($text) -and
            $text -ne '(ConnectData)' -and
            -not $text.StartsWith('#')) {
            return $true
        }
        Start-Sleep -Milliseconds 300
    }
    Write-Host ("      {0} = '{1}' (empty/placeholder after {2}s)" -f $Address, $text, $TimeoutSeconds) -ForegroundColor DarkGray
    return $false
}

# ---------------------------------------------------------------- setup ----

Resolve-RedisCli
Write-Host ("Redis CLI : {0} {1}" -f $script:RedisExe, ($script:RedisPrefix -join ' '))
Write-Host ("Redis host: {0}" -f $RedisHost)
Write-Host ("Key prefix: {0}.*" -f $KeyPrefix)

if ((Invoke-RedisCli @('PING') | Out-String).Trim() -ne 'PONG') {
    throw "Redis is not responding at $RedisHost"
}

$isLocalHost = $RedisHost -match '^(localhost|127\.0\.0\.1)(:\d+)?$'
$allowClientKill = (-not $SkipClientKill) -and $isLocalHost

$script:Excel = New-Object -ComObject Excel.Application
$script:Excel.Visible = $false
$script:Excel.DisplayAlerts = $false
$script:Excel.ScreenUpdating = $false

$udf = $null
$rtd = $null

try {
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
        @{ Row = 12; Func = 'ChannelPublish'; Fx = '=RedisUDFChannelPublish("{1}.channel","ola_mundo","{0}")' -f $h, $kp;                                  Expected = '1 readers(s)' },
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
        @{ Row = 27; Func = 'Get locale';     Fx = '=RedisUDFGet("{1}.locale","{0}")' -f $h, $kp;                                                    Expected = '67000.5' }
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
        @{ Row = 16; Func = 'RealTimeUpdates';     Fx = '=RedisRTDRealTimeUpdates()';                                  Expected = $null }
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
    Start-Sleep -Milliseconds 700
    Invoke-ExcelAction { $script:Excel.CalculateFull() } | Out-Null

    Check (Wait-CellText $udf 'B4' 'OK')                     'UDF Set returns OK'
    Check (Wait-CellText $udf 'B5' 'hello_from_udf')         'UDF Get returns the value'
    Check (Wait-CellText $udf 'B6' '1')                      'UDF Exists returns 1'
    Check (Wait-CellText $udf 'B7' '-1')                     'UDF TTL returns -1 (no expiry)'
    Check (Wait-CellText $udf 'B8' 'OK')                     'UDF SetJSON returns OK'
    Check (Wait-CellText $udf 'B9' '2')                      'UDF JSONToMatrix index [2,1] is 2'
    Check (Wait-CellText $udf 'B10' 'OK')                    'UDF HashSet returns OK'
    Check (Wait-CellText $udf 'B11' 'valor1')                'UDF HashGet returns the value'
    Check (Wait-CellText $udf 'B13' 'ola_mundo')             'UDF ChannelLatest received the published message'
    Check ((Get-CellNumber $udf 'B14') -ge 1)                'UDF ConnectionCount >= 1'
    Check (Wait-CellText $udf 'B15' '1')                     'UDF ExistsMultiples (pipelined) first key exists'
    Check (Wait-CellText $udf 'B16' '-1')                    'UDF TTLMultiples (pipelined) returns -1'
    Check (Wait-CellText $udf 'B17' 'valor1')                'UDF HashGetFieldMultipleKeys (pipelined) returns the value'
    Check (Wait-CellText $udf 'B18' 'OK')                    'UDF SetEx returns OK'
    $ttlOk = $false
    $ttlLast = ''
    $ttlDeadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $ttlDeadline) {
        $ttlValue = 0.0
        $ttlLast = (Get-CellText $udf 'B19').Replace(',', '.')
        if ([double]::TryParse($ttlLast, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ttlValue) -and $ttlValue -gt 0 -and $ttlValue -le 100) { $ttlOk = $true; break }
        Start-Sleep -Milliseconds 300
    }
    if (-not $ttlOk) { Write-Host ("      B19 = '{0}'" -f $ttlLast) -ForegroundColor DarkGray }
    Check $ttlOk 'UDF TTL sees the SetEx expiry'
    Check (Wait-CellText $udf 'B20' '0')                     'UDF Del reports 0 for a missing key'
    Check ((Get-CellText $udf 'B21') -match '^\d+$')         'UDF Incr returns an integer'
    Check ((Get-CellText $udf 'B22') -match '^\d+$')         'UDF ListPushRight returns an integer'
    Check (Wait-CellText $udf 'B23' 'a')                     'UDF ListRange returns the first element'
    Check ((Get-CellText $udf 'B24') -match '^\d+$')         'UDF SetAdd returns an integer'
    Check (Wait-CellText $udf 'B25' 'x')                     'UDF SetMembers returns the member'
    Check (Wait-CellText $udf 'B26' 'OK')                        'UDF Set stores numeric cells invariantly'
    Check (Wait-CellText $udf 'B27' '67000.5')                   'UDF Get returns the invariant number'

    Check (Wait-CellText $rtd 'B4' 'hello_from_udf')         'RTD GET returns the value'
    Check (Wait-CellText $rtd 'B5' 'valor1')                 'RTD HGET returns the value'
    Check (Wait-CellText $rtd 'B6' '{"campo1":"valor1"}')    'RTD HGETALL returns valid JSON'

    Invoke-RedisCli @('PUBLISH', "$KeyPrefix.rtd", 'ALTA') | Out-Null
    Check (Wait-CellText $rtd 'B7' 'ALTA')                   'RTD SUB received the published message'
    Check (Wait-CellText $rtd 'B8' 'ALTA')                   'RTD PSUB received the published message'
    Check ((Get-CellNumber $rtd 'B9') -ge 1)                 'RTD ConnectionCount >= 1'
    Check ((Get-CellNumber $rtd 'B10') -ge 5)                'RTD TopicCount >= 5'
    Check ((Get-CellNumber $rtd 'B11') -ge 2)                'RTD SubscriptionCount >= 2'
    Check ((Get-CellNumber $rtd 'B12') -ge 1)                'RTD ChannelCount >= 1'

    if ($RealChannel) {
        Invoke-RedisCli @('PUBSUB', 'NUMSUB', $RealChannel) | Out-Null
        Check (Wait-CellNotEmpty $rtd 'B17' 30)              'RTD SUB received live data from the real channel'
    }
    if ($RealPattern) {
        Check (Wait-CellNotEmpty $rtd 'B18' 30)              'RTD PSUB received live data from the real pattern'
    }

    # ------------------------------------------------------------- save it ----
    $outPath = Join-Path $env:TEMP 'RedisExcel.Test.xlsx'
    if ($isLocalHost) {
        $outDir = Join-Path $RepoRoot 'test'
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
        $outPath = Join-Path $outDir 'RedisExcel.Test.xlsx'
    }
    else {
        Write-Host "Remote host: the workbook will not be saved into the repository." -ForegroundColor DarkGray
    }
    Remove-Item $outPath -Force -ErrorAction SilentlyContinue
    Invoke-ExcelAction { $script:Workbook.SaveAs($outPath, 51) } | Out-Null
    Check (Test-Path $outPath) ("test workbook saved to " + $outPath)

    # ------------------------------------ regression: copy of the workbook ----
    # The reported bug: copying/closing a workbook silently killed the Pub/Sub
    # subscriptions of the other workbook using the same channel.
    $copyPath = Join-Path $env:TEMP 'RedisExcel.Test.Copy.xlsx'
    Remove-Item $copyPath -Force -ErrorAction SilentlyContinue
    Invoke-ExcelAction { $script:Workbook.SaveCopyAs($copyPath) } | Out-Null
    $copy = Invoke-ExcelAction { $script:Excel.Workbooks.Open($copyPath) }
    $copyRtd = $copy.Worksheets.Item('RTD')

    Invoke-RedisCli @('PUBLISH', "$KeyPrefix.rtd", 'copy-1') | Out-Null
    Check (Wait-CellText $rtd 'B7' 'copy-1')                 'original received copy-1'
    Check (Wait-CellText $copyRtd 'B7' 'copy-1')             'copy received copy-1'

    # Duplicate the RTD sheet inside the copy: new topics for the same channel.
    Invoke-ExcelAction { $copyRtd.Copy($copy.Worksheets.Item($copy.Worksheets.Count)) } | Out-Null
    $dup = $script:Excel.ActiveSheet
    if ($dup.Name -eq 'RTD') { $dup = $copy.Worksheets.Item(2) }
    Invoke-RedisCli @('PUBLISH', "$KeyPrefix.rtd", 'copy-2') | Out-Null
    Check (Wait-CellText $rtd 'B7' 'copy-2')                 'original keeps receiving after a sheet is copied'
    Check (Wait-CellText $copyRtd 'B7' 'copy-2')             'duplicated workbook keeps receiving'
    Check (Wait-CellText $dup 'B7' 'copy-2')                 'duplicated sheet receives'

    Invoke-ExcelAction { $copy.Close($false) } | Out-Null
    Start-Sleep -Seconds 1
    Invoke-RedisCli @('PUBLISH', "$KeyPrefix.rtd", 'copy-3') | Out-Null
    Check (Wait-CellText $rtd 'B7' 'copy-3')                 'original keeps receiving after the copy workbook is closed'

    # ---------------------------------------- regression: connection blip ----
    if ($allowClientKill) {
        Invoke-RedisCli @('CLIENT', 'KILL', 'TYPE', 'pubsub') | Out-Null
        Start-Sleep -Seconds 2
        Invoke-RedisCli @('PUBLISH', "$KeyPrefix.rtd", 'reconnect-1') | Out-Null
        Check (Wait-CellText $rtd 'B7' 'reconnect-1' 30)     'subscriptions recover after the Pub/Sub connection is killed'
    }
    else {
        Write-Host "SKIP  CLIENT KILL (not a local host or -SkipClientKill); subscriptions not tested against a blip." -ForegroundColor DarkGray
    }
}
finally {
    try { if ($copy) { Invoke-ExcelAction { $copy.Close($false) } | Out-Null } } catch { }
    try { if ($script:Workbook) { Invoke-ExcelAction { $script:Workbook.Close($false) } | Out-Null } } catch { }
    if (-not $KeepExcelOpen -and $script:Excel -ne $null) {
        try { Invoke-ExcelAction { $script:Excel.Quit() } | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Excel) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if ($script:Failures -eq 0) {
    Write-Host 'ALL PASS' -ForegroundColor Green
    exit 0
}
Write-Host ("{0} FAILURE(S)" -f $script:Failures) -ForegroundColor Red
exit 1
