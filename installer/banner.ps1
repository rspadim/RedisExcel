<#
  banner.ps1 - the console header RedisExcel prints before its own output.

  Why: the installer runs these helpers in a visible console (setting up a
  server, testing the connection, opening the sample sheet). A window that
  appears, prints terse technical lines and vanishes is exactly what a user
  reads as suspicious - the single most common "is this a virus?" reaction. A
  clear, self-identifying header answers "what is this, who makes it, where did
  it come from" before any command runs.

  Usage (dot-source it, then print):
    . "$PSScriptRoot\banner.ps1"
    Write-RedisExcelBanner -Action 'Checking the Redis connection'

  Output is ASCII-only ON PURPOSE: the helpers run in consoles whose codepage is
  the OEM one (437/850/866/932...), where box-drawing or accented characters
  turn into mojibake. ASCII renders identically everywhere.

  The art is skipped in a window narrower than MIN_WIDTH (or with no console at
  all, e.g. output redirected to a file); a compact one-line header is printed
  instead, so a narrow CONSOLE never wraps the art into a mess.
#>

$script:RedisExcelBannerHome = 'https://github.com/rspadim/RedisExcel'
$script:RedisExcelBannerMinWidth = 70

# 5-row "standard" style art spelling "Redis Excel", one string per row.
# Baseline on the last row; lowercase sits on the x-height, ascenders (d, l)
# reach the first row. Kept as plain data: no font engine, no dependency.
$script:RedisExcelArt = @(
  ' ____              _  _           _____                     _ '
  '|  _ \   ___    __| |(_)  ___    | ____|__  __  ___   ___  | |'
  '| |_) | / _ \  / _` || | / __|   |  _|  \ \/ / / __| / _ \ | |'
  '|  _ < |  __/ | (_| || | \__ \   | |___  >  < | (__ |  __/ | |'
  '|_| \_\ \___|  \__,_||_| |___/   |_____|/_/\_\ \___| \___| |_|'
)

function Get-RedisExcelVersion {
  param([string]$InstallDir = '')

  # 1) The installed DLL carries ProductVersion (e.g. "v1.6.4+<commit>").
  $candidates = @()
  if ($InstallDir -ne '') { $candidates += (Join-Path $InstallDir 'RedisExcel.dll') }
  try {
    # 2) The helper's own folder and its parent ({app}\scripts -> {app}).
    $candidates += (Join-Path $PSScriptRoot 'RedisExcel.dll')
    $candidates += (Join-Path (Split-Path -Parent $PSScriptRoot) 'RedisExcel.dll')
  } catch { }

  foreach ($dll in $candidates) {
    try {
      if ($dll -and (Test-Path -LiteralPath $dll)) {
        $v = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dll).ProductVersion
        if ($v) {
          # Local builds report "v1.6.4+<sha>"; the sha is noise in a header.
          $plus = $v.IndexOf('+')
          if ($plus -gt 0) { $v = $v.Substring(0, $plus) }
          return $v.Trim()
        }
      }
    } catch { }
  }
  return ''
}

function Test-BannerWideEnough {
  param([int]$MinWidth = 70)
  try {
    $width = [Console]::WindowWidth
    if (-not $width -or $width -lt 1) { return $false }
    return ($width -ge $MinWidth)
  } catch {
    # A redirected stream throws here: assume compact output.
    return $false
  }
}

function Write-RedisExcelBanner {
  param(
    # Short line describing what THIS run does ("Installing the Redis server").
    [string]$Action = '',
    # Installed folder, when the caller knows it.
    [string]$InstallDir = '',
    # Force the full art even when the window looks narrow.
    [switch]$Force
  )

  $version = Get-RedisExcelVersion -InstallDir $InstallDir
  $tagline = 'Redis + Excel add-in (XLL)'
  if ($version) { $tagline = $tagline + '  -  ' + $version }

  $wide = $true
  if (-not $Force) { $wide = Test-BannerWideEnough -MinWidth $script:RedisExcelBannerMinWidth }

  Write-Host ''
  if ($wide) {
    $art = $script:RedisExcelArt
    for ($r = 0; $r -lt $art.Count; $r++) {
      Write-Host ('  ' + $art[$r]) -ForegroundColor Red
    }
    Write-Host ''
    Write-Host ('  ' + $tagline) -ForegroundColor Gray
    Write-Host ('  ' + $script:RedisExcelBannerHome) -ForegroundColor DarkGray
  } else {
    # Compact: the product name already carries the version, so the tagline
    # must NOT repeat it (it read "RedisExcel v1.6.4 - ... - v1.6.4").
    $name = 'RedisExcel'
    if ($version) { $name = $name + ' ' + $version }
    Write-Host ('  ' + $name + ' - Redis + Excel add-in (XLL)') -ForegroundColor Red
    Write-Host ('  ' + $script:RedisExcelBannerHome) -ForegroundColor DarkGray
  }

  if ($Action -ne '') {
    Write-Host ''
    Write-Host ('  ' + $Action) -ForegroundColor White
  }
  if ($InstallDir -ne '') {
    Write-Host ('  Folder: ' + $InstallDir) -ForegroundColor DarkGray
  }
  Write-Host ''
}
