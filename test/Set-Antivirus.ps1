<#
.SYNOPSIS
    Report and (where the platform allows) toggle antivirus real-time
    protection, so the Excel E2E can be run with protection ON or OFF.

.DESCRIPTION
    Excel COM automation is intermittently blocked by security software
    (see AGENTS.md, "Excel automation lessons"). This helper is intentionally
    VENDOR-AGNOSTIC: it discovers whatever is registered with the Windows
    Security Center and only uses mechanisms that work across products:

      * Status  - lists every registered AV product (state nibble decoded) and
                  every service that looks like an AV engine.
      * Disable - turns off Microsoft Defender real-time protection through the
                  documented Set-MpPreference cmdlet, and best-effort stops the
                  AV services it can (running as admin).
      * Enable  - reverses the above (re-enables Defender real-time, starts the
                  services it stopped).

    IMPORTANT LIMIT (by design, not a bug): a third-party engine with its
    "self-defense"/tamper-protection enabled cannot be stopped by ANY generic
    API - the service refuses to stop and the process cannot be killed, and the
    home products expose no non-interactive "disable protection" switch. For
    those, Disable stops whatever it can and tells the operator to pause the
    engine from its own UI/tray. The harness therefore must stay correct with
    protection ON and only be compared with it OFF as an operator action; use
    Status to record which state a run had.

.PARAMETER Action
    Status (default), Disable or Enable.

.PARAMETER WhatIfOnly
    Show what Disable/Enable would do without changing anything.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File test\Set-Antivirus.ps1 -Action Status
    powershell -ExecutionPolicy Bypass -File test\Set-Antivirus.ps1 -Action Disable
    powershell -ExecutionPolicy Bypass -File test\Set-Antivirus.ps1 -Action Enable
#>
param(
    [ValidateSet('Status', 'Disable', 'Enable')]
    [string]$Action = 'Status',
    [switch]$WhatIfOnly
)

$ErrorActionPreference = 'Stop'

function Get-AvProducts {
    # Windows Security Center is the product-independent source: every
    # registered AV reports here with a productState whose 0x?000 nibble is the
    # enabled/snoozed code (0 = off).
    try {
        Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop |
            ForEach-Object {
                $state = [int]$_.productState
                [pscustomobject]@{
                    Name   = [string]$_.displayName
                    # Real-time state is the 0x1000 bit (0x41000 = enabled,
                    # 0x42000 = paused); testing the bit is product-independent.
                    Active = (($state -band 0x1000) -ne 0)
                    State  = '0x{0:X}' -f $state
                }
            }
    }
    catch {
        @([pscustomobject]@{ Name = "unavailable ($($_.Exception.Message))"; Active = $false; State = '' })
    }
}

function Get-AvServices {
    # Heuristic: services whose name or display name looks like an AV engine.
    # No vendor is hard-coded; this catches mainstream products generically.
    Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName -match '(?i)anti-?v[íi]rus' -or
            $_.Name -match '(?i)^(avp|avast|avgsvc|ekrn|nod32|mcafee|msmpeng|symantec|sophos|windefend|wdnissvc)'
        } |
        Select-Object Name, DisplayName, State, StartMode, ProcessId
}

function Test-Admin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-Status {
    Write-Host 'AV products (Windows Security Center):' -ForegroundColor Cyan
    foreach ($p in Get-AvProducts) {
        $tag = if ($p.Active) { 'ACTIVE' } else { 'inactive' }
        Write-Host ("  {0,-28} {1,-9} {2}" -f $p.Name, $tag, $p.State)
    }
    Write-Host 'AV-looking services:' -ForegroundColor Cyan
    Get-AvServices | ForEach-Object {
        Write-Host ("  {0,-20} {1,-8} {2}" -f $_.Name, $_.State, $_.DisplayName)
    }
}

$defenderPresent = [bool](Get-Module -ListAvailable -Name Defender -ErrorAction SilentlyContinue)
$isAdmin = Test-Admin

switch ($Action) {
    'Status' { Show-Status }
    'Disable' {
        if (-not $isAdmin) { throw "Disable needs an elevated PowerShell (admin)." }
        Show-Status
        if ($defenderPresent) {
            Write-Host 'Disabling Microsoft Defender real-time protection...' -ForegroundColor Yellow
            if (-not $WhatIfOnly) {
                try { Set-MpPreference -DisableRealtimeMonitoring $true -ErrorAction Stop; Write-Host '  Defender real-time: OFF' }
                catch { Write-Host ("  Defender real-time: could not disable ({0})" -f $_.Exception.Message) -ForegroundColor Red }
            }
        }
        foreach ($s in Get-AvServices | Where-Object { $_.State -eq 'Running' }) {
            Write-Host ("Stopping service {0}..." -f $s.Name) -ForegroundColor Yellow
            if (-not $WhatIfOnly) {
                try { Stop-Service -Name $s.Name -Force -ErrorAction Stop; Write-Host '  stopped' }
                catch { Write-Host ("  could not stop ({0}) - self-protected; pause it in its own UI" -f $_.Exception.Message) -ForegroundColor Red }
            }
        }
        Write-Host 'Re-check:' -ForegroundColor Cyan; Show-Status
        Write-Host 'Note: a self-protected third-party engine can only be paused from its own product UI.' -ForegroundColor DarkGray
    }
    'Enable' {
        if (-not $isAdmin) { throw "Enable needs an elevated PowerShell (admin)." }
        if ($defenderPresent) {
            Write-Host 'Re-enabling Microsoft Defender real-time protection...' -ForegroundColor Yellow
            if (-not $WhatIfOnly) {
                try { Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop; Write-Host '  Defender real-time: ON' }
                catch { Write-Host ("  could not enable ({0})" -f $_.Exception.Message) -ForegroundColor Red }
            }
        }
        foreach ($s in Get-AvServices | Where-Object { $_.State -ne 'Running' -and $_.StartMode -ne 'Disabled' }) {
            Write-Host ("Starting service {0}..." -f $s.Name) -ForegroundColor Yellow
            if (-not $WhatIfOnly) {
                try { Start-Service -Name $s.Name -ErrorAction Stop; Write-Host '  started' }
                catch { Write-Host ("  could not start ({0})" -f $_.Exception.Message) -ForegroundColor Red }
            }
        }
        Write-Host 'Re-check:' -ForegroundColor Cyan; Show-Status
    }
}
