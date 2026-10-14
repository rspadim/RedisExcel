<#
  validate-config.ps1 - checks a RedisExcel.json against the SAME schema the
  add-in uses (RedisExcel.schema.json, shipped next to this script).

  Very verbose on purpose: it prints every key, every problem and a clear
  OK/FAIL summary, so the user is never in doubt.

  Usage:
    powershell -ExecutionPolicy Bypass -File validate-config.ps1
    powershell -ExecutionPolicy Bypass -File validate-config.ps1 -Config "C:\RedisExcel\RedisExcel.json"
    powershell -ExecutionPolicy Bypass -File validate-config.ps1 -Schema "C:\RedisExcel\RedisExcel.schema.json"

  It needs no add-in and no Redis: it only reads JSON files.

  Exit codes: 0 = the config is valid (or no config found -> defaults),
              1 = problems found, 2 = the file could not be read/parsed.

  Requires PowerShell 5+ (uses ConvertFrom-Json; a tiny built-in validator
  covers the types, ranges, enums and the pattern rules of the schema).
#>
[CmdletBinding()]
param(
  [string]$Config = '',
  [string]$Schema = ''
)

$ErrorActionPreference = 'Stop'

function Line([string]$m = '') { Write-Host $m }

Line '======================================================================'
Line ' RedisExcel - config validator'
Line '======================================================================'

# ---- Locate the files ------------------------------------------------------
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Schema) {
  $candidates = @(
    (Join-Path $here 'RedisExcel.schema.json'),
    (Join-Path (Split-Path -Parent $here) 'RedisExcel.schema.json')
  )
  $Schema = ($candidates | Where-Object { Test-Path $_ } | Select-Object -First 1)
}
if (-not $Schema -or -not (Test-Path $Schema)) {
  Line 'FAIL: RedisExcel.schema.json was not found next to this tool.'
  Line '      Run this tool from the folder where RedisExcel is installed.'
  exit 2
}
Line ("Schema : " + (Resolve-Path $Schema).Path)

if (-not $Config) {
  # Same search order the add-in uses, plus the folder of this tool and its parent.
  $excelDir = Split-Path -Parent (Get-Process -Id $PID).MainModule.FileName
  $parent  = Split-Path -Parent $here
  $candidates = @(
    (Join-Path $here 'RedisExcel.json'),
    (Join-Path $parent 'RedisExcel.json'),
    (Join-Path $env:USERPROFILE 'RedisExcel.json'),
    (Join-Path $excelDir 'RedisExcel.json'),
    (Join-Path $env:WINDIR 'RedisExcel.json')
  )
  $Config = ($candidates | Where-Object { Test-Path $_ } | Select-Object -First 1)
}
if (-not $Config -or -not (Test-Path $Config)) {
  Line 'Config : (none found)'
  Line ''
  Line 'No RedisExcel.json was found. The add-in will use its built-in defaults'
  Line '(host localhost:6379). This is OK.'
  Line ''
  Line 'RESULT: OK (no config file; defaults will be used)'
  exit 0
}
Line ("Config : " + (Resolve-Path $Config).Path)
Line ''

# ---- Parse (accept // and /* */ comments like the add-in) ------------------
$raw = Get-Content -LiteralPath $Config -Raw

# Remove comments AND trailing commas WITHOUT touching either inside strings (so
# "redis://host" and "a,}" stay intact). The add-in uses Newtonsoft, which
# accepts both; the PowerShell parser does not, so the validator normalizes them
# or it would flag a file the add-in loads fine (and a naive regex corrupted a
# string value, audit G1/R7).
function Normalize-Json([string]$text) {
  $sb = New-Object System.Text.StringBuilder
  $inString = $false
  $inLine = $false
  $inBlock = $false
  $i = 0
  $n = $text.Length
  while ($i -lt $n) {
    $c = $text[$i]
    $c2 = if ($i + 1 -lt $n) { $text[$i + 1] } else { [char]0 }
    if ($inLine) {
      if ($c -eq "`n") { $inLine = $false; [void]$sb.Append($c) }
      $i++
      continue
    }
    if ($inBlock) {
      if ($c -eq '*' -and $c2 -eq '/') { $inBlock = $false; $i += 2 } else { $i++ }
      continue
    }
    if ($inString) {
      [void]$sb.Append($c)
      if ($c -eq '\') { if ($i + 1 -lt $n) { [void]$sb.Append($c2); $i += 2; continue } }
      elseif ($c -eq '"') { $inString = $false }
      $i++
      continue
    }
    if ($c -eq '"') { $inString = $true; [void]$sb.Append($c); $i++; continue }
    if ($c -eq '/' -and $c2 -eq '/') { $inLine = $true; $i += 2; continue }
    if ($c -eq '/' -and $c2 -eq '*') { $inBlock = $true; $i += 2; continue }
    if ($c -eq ',') {
      # Drop a comma whose next non-whitespace char closes an object/array.
      $j = $i + 1
      while ($j -lt $n -and ([char]$text[$j] -eq ' ' -or [char]$text[$j] -eq "`t" -or [char]$text[$j] -eq "`r" -or [char]$text[$j] -eq "`n")) { $j++ }
      if ($j -lt $n -and ([char]$text[$j] -eq '}' -or [char]$text[$j] -eq ']')) { $i++; continue }
    }
    [void]$sb.Append($c)
    $i++
  }
  return $sb.ToString()
}

# Best-effort: true when two property names in the document differ only by case
# (the add-in merges these with a warning; PowerShell refuses to parse them).
function Test-HasCaseVariantDuplicateKey([string]$text) {
  $seen = @{}
  foreach ($m in [regex]::Matches($text, '"((?:[^"\\]|\\.)*)"\s*:')) {
    $name = $m.Groups[1].Value
    $key = $name.ToLowerInvariant()
    if ($seen.ContainsKey($key) -and $seen[$key] -cne $name) { return $true }
    $seen[$key] = $name
  }
  return $false
}

$json = $null
$duplicateKeys = $false
$parseError = $null
try {
  $json = $raw | ConvertFrom-Json
} catch {
  $stripped = Normalize-Json $raw
  try { $json = $stripped | ConvertFrom-Json }
  catch {
    # PowerShell's ConvertFrom-Json rejects duplicate keys (case-variant too),
    # but the add-in accepts them (a warning). Do not report a hard FAIL for a
    # file the add-in loads. The exception text is localized, so detect the
    # duplicate case by scanning the object keys instead of the message wording.
    if (Test-HasCaseVariantDuplicateKey $stripped) { $duplicateKeys = $true }
    else { $parseError = $_.Exception.Message }
  }
}

if ($duplicateKeys) {
  Line 'WARNING: the file has duplicate keys (PowerShell cannot parse them).'
  Line '         The add-in accepts case-variant duplicates with a warning; remove the'
  Line '         duplicates for a full check (e.g. "Default" and "default").'
  Line 'RESULT: OK (with 1 warning: duplicate keys) - not fully validated.'
  exit 0
}
if ($null -eq $json) {
  # The file IS JSON the add-in can read, but PowerShell could not parse it (a
  # dialect the add-in tolerates). Do not claim it is invalid.
  Line 'FAIL: the file is not valid JSON.'
  if ($parseError) { Line ("      " + $parseError) }
  exit 2
}
# The root must be a JSON OBJECT: the add-in rejects anything else (empty file,
# array, number, string, bool, null, comment-only) with #CONFIG - the validator
# must agree instead of reporting OK (audit F3). A scalar/array root parses as a
# non-PSCustomObject, and an empty/comment-only file parses as $null.
if ($null -eq $json -or -not ($json -is [System.Management.Automation.PSCustomObject])) {
  Line 'FAIL: the config root is not a JSON object.'
  Line '      The add-in requires a JSON object ({ ... }); anything else fails closed (#CONFIG).'
  exit 2
}
Line 'The file parsed as JSON. Checking each key against the schema...'
Line ''

$schemaObj = Get-Content -LiteralPath $Schema -Raw | ConvertFrom-Json

# ---- Tiny schema walker (types, ranges, enums, patterns) -------------------
# $problems are ERRORS (the add-in rejects the whole file -> #CONFIG);
# $warnings are tolerated (extra keys, logged only).
$problems = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]
$checked  = 0

function TypeOk($value, $type) {
  switch ($type) {
    'string'  { return ($value -is [string]) }
    'integer' { return ($value -is [int] -or $value -is [long]) }
    'number'  { return ($value -is [int] -or $value -is [long] -or $value -is [double]) }
    'boolean' { return ($value -is [bool]) }
    'object'  { return ($value -is [System.Management.Automation.PSCustomObject]) }
    'array'   { return ($value -is [System.Array]) }
    default   { return $true }
  }
}

function Describe($value) {
  if ($null -eq $value) { return 'null' }
  return ($value.GetType().Name + " '" + ($value | Out-String).Trim() + "'")
}

function CheckValue($path, $value, $sub, [switch]$Strict) {
  $script:checked++
  # oneOf (e.g. a Servers value: a connection string OR an object). Valid when at
  # least one branch validates; otherwise it is an ERROR, matching the add-in's
  # fail-closed rejection (the previous walker ignored oneOf and reported a false
  # OK for a numeric/array/bad-object Servers value - audit E2). Inside a branch an
  # unknown key is a branch failure (the branch declares additionalProperties:false),
  # not the tolerant top-level warning.
  if ($sub.oneOf) {
    $ok = $false
    foreach ($branch in $sub.oneOf) {
      $probe = New-Object System.Collections.Generic.List[string]
      $saved = $script:problems
      $script:problems = $probe
      try { $r = CheckValue $path $value $branch -Strict } finally { $script:problems = $saved }
      if ($r) { $ok = $true; break }
    }
    if (-not $ok) {
      $script:problems.Add("$path : expected a connection string or an object with host/connectTimeout/syncTimeout")
      return $false
    }
    return $true
  }
  $typeOk = $true
  if ($sub.type) {
    if (-not (TypeOk $value $sub.type)) {
      $script:problems.Add("$path : expected " + $sub.type + ' but got ' + (Describe $value))
      return $false
    }
  }
  if ($sub.type -eq 'string' -and $value -is [string]) {
    if ($sub.pattern) {
      # The schema's patterns are meant to be case-insensitive ASCII.
      if ($value -notmatch ('^(?i:' + ($sub.pattern -replace '^\(\?i\)', '') + ')$')) {
        $script:problems.Add("$path : '$value' is not one of the allowed values")
        $typeOk = $false
      }
    }
  }
  if ($sub.type -eq 'integer' -or $sub.type -eq 'number') {
    if ($null -ne $sub.minimum -and $value -lt $sub.minimum) {
      $script:problems.Add("$path : $value is below the minimum $($sub.minimum)")
      $typeOk = $false
    }
    if ($null -ne $sub.maximum -and $value -gt $sub.maximum) {
      $script:problems.Add("$path : $value is above the maximum $($sub.maximum)")
      $typeOk = $false
    }
  }
  # Recurse into objects.
  if ($sub.type -eq 'object' -and ($value -is [System.Management.Automation.PSCustomObject])) {
    # additionalProperties semantics: absent = allowed; $false = not allowed;
    # an object = validate the value against that schema; $true = allowed.
    # (A boolean $false is NOT null, so the old `$null -ne` test wrongly treated
    # `additionalProperties:false` as permissive - audit E2.)
    $addSchema = $null
    $addAllowed = $true
    if ($sub.PSObject.Properties.Name -contains 'additionalProperties') {
      $add = $sub.additionalProperties
      if ($add -is [System.Management.Automation.PSCustomObject]) { $addSchema = $add }
      elseif ($add -eq $false) { $addAllowed = $false }
    }
    $known = @()
    if ($sub.properties) { $known = @($sub.properties.PSObject.Properties.Name) }
    # required: every listed property must be present (e.g. Servers object -> host).
    if ($sub.required) {
      $present = @($value.PSObject.Properties.Name)
      foreach ($req in $sub.required) {
        if ($present -notcontains $req) {
          $script:problems.Add("$path : missing required property '$req'")
          $typeOk = $false
        }
      }
    }
    foreach ($prop in $value.PSObject.Properties) {
      $name = $prop.Name
      if ($sub.properties -and ($known -contains $name)) {
        if (-not (CheckValue ("$path/" + $name) $prop.Value $sub.properties.$name)) { $typeOk = $false }
      } elseif ($null -ne $addSchema) {
        # Free-form map (e.g. Servers): validate the value against the
        # additionalProperties schema (oneOf in the Servers case).
        if (-not (CheckValue ("$path/" + $name) $prop.Value $addSchema)) { $typeOk = $false }
      } elseif ($addAllowed) {
        # additionalProperties: true -> anything is allowed.
      } elseif ($Strict) {
        # Inside a oneOf branch: an unknown key makes THIS branch fail (the branch
        # declares additionalProperties:false), it is not a tolerant warning.
        $typeOk = $false
      } else {
        # An extra/unknown key is only a WARNING (the add-in logs and ignores
        # it); it does NOT reject the file.
        $script:warnings.Add("$path/$name : unknown key (ignored by the add-in)")
      }
    }
  }
  return $typeOk
}

# Walk top-level keys.
$topSchema = $schemaObj
$topProps = @($topSchema.properties.PSObject.Properties.Name)

# The removed v1 keys are ERRORS in the add-in (fail-closed): mirror that so the
# validator agrees with the loader.
if ($json.PSObject.Properties.Name -contains 'UDF') {
  $problems.Add("/UDF : removed in v1.6; declare servers under `"Servers`" and pick one with `"Default`"")
}
if ($json.PSObject.Properties.Name -contains 'CoalesceRealtimeUpdates') {
  $problems.Add("/CoalesceRealtimeUpdates : moved into `"RTD`" in v1.6")
}
if ($json.RTD -and ($json.RTD.PSObject.Properties.Name -contains 'host')) {
  $problems.Add("/RTD/host : removed in v1.6; use `"Default`" (or a `"Servers`" name)")
}
if ($json.RTD -and ($json.RTD.PSObject.Properties.Name -contains 'timeout')) {
  $problems.Add("/RTD/timeout : removed in v1.6; use the server's connectTimeout/syncTimeout or the global `"Timeout`"")
}

# A dangling Default (a bare word not in Servers and not a connection string) is
# only a WARNING in the add-in: a single-label host ("localhost", a service name)
# is a valid port-less connection string, so rejecting it would lock a valid
# config out (audit F1). A real host is used as-is, never a silent fallback.
$defaultName = $null
if ($json.Default -is [string]) { $defaultName = $json.Default.Trim() }
if ($defaultName) {
  $looksLikeConn = ($defaultName -match '[:,=/]') -or ($defaultName -match '\.') -or ($defaultName -ieq 'localhost')
  $inServers = $false
  if ($json.Servers) {
    foreach ($p in $json.Servers.PSObject.Properties) {
      if ($p.Name -eq $defaultName -or $p.Name.ToLowerInvariant() -eq $defaultName.ToLowerInvariant()) { $inServers = $true; break }
    }
  }
  if (-not $looksLikeConn -and -not $inServers) {
    $warnings.Add("/Default : '$defaultName' is not defined in `"Servers`" and is not a connection string; using it as a host")
  }
}

foreach ($name in $json.PSObject.Properties.Name) {
  $value = $json.$name
  if ($topProps -notcontains $name) {
    $warnings.Add("/$name : unknown key (ignored by the add-in)")
    continue
  }
  [void](CheckValue ("/" + $name) $value $topSchema.properties.$name)
  Line ("  checked  /$name")
}

# Connection-string sanity the JSON schema cannot express: StackExchange.Redis
# splits options on commas, so a `password=` value that continues after a comma
# is SILENTLY TRUNCATED (password=a,b -> the client uses "a" and reads "b" as
# another option; there is no quoting that works and the URI form carries no
# credentials). Warn on the shape that looks like a truncated password: a
# comma-separated segment after password= that is neither key=value nor host:port.
function Test-TruncatedPassword([string]$conn) {
  $i = $conn.IndexOf('password=', [System.StringComparison]::OrdinalIgnoreCase)
  if ($i -lt 0) { return $false }
  $rest = $conn.Substring($i + 9)
  $segments = $rest -split ','
  for ($k = 1; $k -lt $segments.Count; $k++) {
    $s = $segments[$k].Trim()
    if ($s -eq '') { continue }
    if ($s -match '=') { continue }                       # key=value option
    # A host segment: a name/ip WITH a port, a dotted name, localhost, or a
    # bracketed IPv6. A bare word (no port/dot) is not treated as a host: it is
    # the signature of a password that continued after the comma.
    if ($s -match '^\[[0-9A-Fa-f:]+\](:\d+)?$') { continue }   # [::1]:6379
    if ($s -match '^[A-Za-z0-9\-\.]*[A-Za-z\-][A-Za-z0-9\-\.]*:\d+$') { continue } # name:port
    if ($s -match '^([A-Za-z0-9\-]+\.)+[A-Za-z0-9\-]+(:\d+)?$') { continue }  # dotted name
    if ($s -match '^localhost(:\d+)?$') { continue }
    return $true
  }
  return $false
}
function Check-ConnectionStrings($node, [string]$path) {
  if ($null -eq $node) { return }
  if ($node -is [string]) {
    if (Test-TruncatedPassword $node) {
      $warnings.Add("$path : the password looks truncated - a comma after password= starts a new option (use a password without commas)")
    }
    return
  }
  if ($node -is [System.Management.Automation.PSCustomObject]) {
    foreach ($p in $node.PSObject.Properties) { Check-ConnectionStrings $p.Value ("$path/" + $p.Name) }
    return
  }
  if ($node -is [System.Collections.IEnumerable] -and -not ($node -is [string])) {
    $n = 0
    foreach ($item in $node) { Check-ConnectionStrings $item ("$path[" + $n + "]"); $n++ }
  }
}
Check-ConnectionStrings $json.Default '/Default'
if ($json.Servers) { Check-ConnectionStrings $json.Servers '/Servers' }

Line ''
Line '----------------------------------------------------------------------'
if ($problems.Count -eq 0) {
  if ($warnings.Count -eq 0) {
    Line ("RESULT: OK - " + $checked + ' value(s) checked, no problems.')
  } else {
    Line ("RESULT: OK - " + $checked + ' value(s) checked, ' + $warnings.Count + ' warning(s).')
    foreach ($w in $warnings) { Line ("  warning: " + $w) }
  }
  Line 'Your RedisExcel.json is valid. Restart Excel if you just edited it.'
  exit 0
} else {
  Line ("RESULT: FAIL - " + $problems.Count + ' problem(s) in ' + $checked + ' checked value(s):')
  foreach ($p in $problems) { Line ("  - " + $p) }
  if ($warnings.Count -gt 0) {
    Line ''
    Line ("  " + $warnings.Count + ' warning(s) (not fatal):')
    foreach ($w in $warnings) { Line ("  warning: " + $w) }
  }
  Line ''
  Line 'A wrong VALUE on a known key makes the add-in REJECT the whole file:'
  Line 'every cell shows "#CONFIG" until it is fixed (there is no silent'
  Line 'fallback). Extra/unknown keys are only warnings.'
  exit 1
}
