<#
.SYNOPSIS
    Compares two Excel E2E matrix runs (test\Run-ExcelE2E-Matrix.ps1 -Tag ...)
    side by side, so the antivirus on/off impact is visible per mode.

.DESCRIPTION
    Reads %TEMP%\e2e-<tag>-<mode>.log for the four modes and prints the result,
    the COM-retry total (the AV fingerprint) and the SKIP count. Use the same
    script for both tags so the only variable is the environment.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File test\Compare-ExcelE2E.ps1 -TagA avon -TagB avoff
#>
param(
    [Parameter(Mandatory = $true)][string]$TagA,
    [Parameter(Mandatory = $true)][string]$TagB
)

function Read-MatrixOne([string]$tag) {
    $modes = [ordered]@{}
    foreach ($mode in 'default', 'sync', 'ffall', 'async') {
        $log = Join-Path $env:TEMP ("e2e-{0}-{1}.log" -f $tag, $mode)
        if (-not (Test-Path -LiteralPath $log)) { continue }
        $text = Get-Content -LiteralPath $log -Raw
        $status = if ($text -match 'ALL PASS') { 'ALL PASS' }
                  elseif ($text -match '(\d+) FAILURE') { "FAIL x$($Matches[1])" }
                  else { 'incomplete' }
        $null = $text -match 'COM retries: total=(\d+), worst-single=(\d+)'
        $retries = if ($Matches -and $Matches.Count -ge 3) { [int]$Matches[1] } else { -1 }
        $worst = if ($Matches -and $Matches.Count -ge 3) { [int]$Matches[2] } else { -1 }
        $skips = ([regex]::Matches($text, '(?m)^SKIP ')).Count
        $modes[$mode] = [pscustomobject]@{ Status = $status; Retries = $retries; Worst = $worst; Skips = $skips }
    }
    return $modes
}

$a = Read-MatrixOne $TagA
$b = Read-MatrixOne $TagB

Write-Host ("{0,-8} | {1,-16} {2,-8} {3,-6} | {4,-16} {5,-8} {6,-6}" -f `
    'mode', "$TagA status", 'retries', 'skip', "$TagB status", 'retries', 'skip')
Write-Host ('-' * 82)
foreach ($mode in 'default', 'sync', 'ffall', 'async') {
    $ra = $a[$mode]; $rb = $b[$mode]
    Write-Host ("{0,-8} | {1,-16} {2,-8} {3,-6} | {4,-16} {5,-8} {6,-6}" -f `
        $mode,
        ($(if ($ra) { $ra.Status } else { '-' })), ($(if ($ra) { $ra.Retries } else { '-' })), ($(if ($ra) { $ra.Skips } else { '-' })),
        ($(if ($rb) { $rb.Status } else { '-' })), ($(if ($rb) { $rb.Retries } else { '-' })), ($(if ($rb) { $rb.Skips } else { '-' })))
}
Write-Host ''
Write-Host "COM retries ~0 in a tag means a clean environment; a high total with mostly-1-retry means the AV scans every COM call. Status/SKIP differences between the tags are the AV's observable impact."
