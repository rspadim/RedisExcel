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

# Remove comments WITHOUT touching // inside strings (so "redis://host" stays).
function Strip-JsonComments([string]$text) {
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
    [void]$sb.Append($c)
    $i++
  }
  return $sb.ToString()
}

$json = $null
try {
  $json = $raw | ConvertFrom-Json
} catch {
  $stripped = Strip-JsonComments $raw
  try { $json = $stripped | ConvertFrom-Json }
  catch {
    Line 'FAIL: the file is not valid JSON.'
    Line ("      " + $_.Exception.Message)
    exit 2
  }
}
Line 'The file parsed as JSON. Checking each key against the schema...'
Line ''

$schemaObj = Get-Content -LiteralPath $Schema -Raw | ConvertFrom-Json

# ---- Tiny schema walker (types, ranges, enums, patterns) -------------------
$problems = New-Object System.Collections.Generic.List[string]
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

function CheckValue($path, $value, $sub) {
  $script:checked++
  if ($sub.type) {
    if (-not (TypeOk $value $sub.type)) {
      $script:problems.Add("$path : expected " + $sub.type + ' but got ' + (Describe $value))
      return
    }
  }
  if ($sub.type -eq 'string' -and $value -is [string]) {
    if ($sub.pattern) {
      # The schema's patterns are meant to be case-insensitive ASCII.
      if ($value -notmatch ('^(?i:' + ($sub.pattern -replace '^\(\?i\)', '') + ')$')) {
        $script:problems.Add("$path : '$value' is not one of the allowed values")
      }
    }
  }
  if ($sub.type -eq 'integer' -or $sub.type -eq 'number') {
    if ($null -ne $sub.minimum -and $value -lt $sub.minimum) {
      $script:problems.Add("$path : $value is below the minimum $($sub.minimum)")
    }
    if ($null -ne $sub.maximum -and $value -gt $sub.maximum) {
      $script:problems.Add("$path : $value is above the maximum $($sub.maximum)")
    }
  }
  # Recurse into objects.
  if ($sub.type -eq 'object' -and ($value -is [System.Management.Automation.PSCustomObject])) {
    $hasAdditional = ($null -ne $sub.additionalProperties)
    $known = @()
    if ($sub.properties) { $known = @($sub.properties.PSObject.Properties.Name) }
    foreach ($prop in $value.PSObject.Properties) {
      $name = $prop.Name
      if ($sub.properties -and ($known -contains $name)) {
        CheckValue ("$path/" + $name) $prop.Value $sub.properties.$name
      } elseif ($hasAdditional -and ($sub.additionalProperties -is [System.Management.Automation.PSCustomObject])) {
        # Free-form map (e.g. Servers): validate the value against the
        # additionalProperties schema.
        CheckValue ("$path/" + $name) $prop.Value $sub.additionalProperties
      } elseif ($hasAdditional) {
        # additionalProperties: true -> anything is allowed.
      } else {
        $script:problems.Add("$path/$name : unknown key")
      }
    }
  }
}

# Walk top-level keys.
$topSchema = $schemaObj
$topProps = @($topSchema.properties.PSObject.Properties.Name)
foreach ($name in $json.PSObject.Properties.Name) {
  $value = $json.$name
  if ($topProps -notcontains $name) {
    $problems.Add("/$name : unknown key")
    continue
  }
  CheckValue ("/" + $name) $value $topSchema.properties.$name
  Line ("  checked  /$name")
}

Line ''
Line '----------------------------------------------------------------------'
if ($problems.Count -eq 0) {
  Line ("RESULT: OK - " + $checked + ' value(s) checked, no problems.')
  Line 'Your RedisExcel.json is valid. Restart Excel if you just edited it.'
  exit 0
} else {
  Line ("RESULT: FAIL - " + $problems.Count + ' problem(s) in ' + $checked + ' checked value(s):')
  foreach ($p in $problems) { Line ("  - " + $p) }
  Line ''
  Line 'Fix the lines above. A wrong value is usually ignored (the add-in'
  Line 'falls back to the default for that key), but it is listed here so you'
  Line 'know exactly what is off.'
  exit 1
}
