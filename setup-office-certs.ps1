<#
One-click setup for Office Add-ins dev certs and WebView loopback.

Usage examples:
  .\setup-office-certs.ps1
  .\setup-office-certs.ps1 -Days 730 -StartDesktop
  .\setup-office-certs.ps1 -StartWeb
#>

[CmdletBinding()]
param(
  [int]$Days = 365,
  [switch]$StartDesktop,
  [switch]$StartWeb
)

function Info($m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }
function Ok($m){ Write-Host "[ OK ] $m" -ForegroundColor Green }
function Warn($m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function Err($m){ Write-Host "[ERR ] $m" -ForegroundColor Red }

# Elevate to admin if needed
function Ensure-Admin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p = New-Object Security.Principal.WindowsPrincipal($id)
  if(-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    Info "Restarting as Administrator..."
    $argLine = $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
      if ($_.Value -is [switch]) { if ($_.Value.IsPresent) { "-$($_.Key)" } }
      else { "-$($_.Key) `"$($_.Value)`"" }
    }
    $argLine = ($argLine | Where-Object { $_ }) -join ' '
    $psi = New-Object System.Diagnostics.ProcessStartInfo "powershell"
    $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $argLine"
    $psi.Verb = "runas"
    try { [Diagnostics.Process]::Start($psi) | Out-Null; exit } catch { Err "Administrator permission is required."; exit 1 }
  }
}
Ensure-Admin

# Ensure NPX exists
if (-not (Get-Command npx -ErrorAction SilentlyContinue)) {
  Err "npx not found. Please install Node.js (includes npm/npx) and retry."
  exit 1
}

$certDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$caCrt   = Join-Path $certDir "ca.crt"
$locCrt  = Join-Path $certDir "localhost.crt"
$locKey  = Join-Path $certDir "localhost.key"

# 1) Uninstall old certs and clear cache
Info "Uninstalling previous dev certs (ignore errors if none)..."
try { npx office-addin-dev-certs uninstall | Out-Null } catch { }
if (Test-Path $certDir) {
  Info "Removing cache folder: $certDir"
  try { Remove-Item -Recurse -Force $certDir } catch { Warn "Remove cache failed (will continue)..." }
}

# 2) Install new certs (no --verbose for compatibility)
Info "Installing new dev certs (valid for $Days days)..."
try {
  npx office-addin-dev-certs install --days $Days | Out-Null
} catch {
  Err ("office-addin-dev-certs failed: " + $_.Exception.Message)
  exit 1
}

# 3) Check files / fallback discovery
if (-not (Test-Path $certDir)) { Err ("Cert folder not found: " + $certDir); exit 1 }

# Primary: localhost.crt
$leafCrt = $null
if (Test-Path $locCrt) { $leafCrt = $locCrt }
else {
  # Fallback 1: 127.0.0.1.crt
  $alt127 = Join-Path $certDir "127.0.0.1.crt"
  if (Test-Path $alt127) { $leafCrt = $alt127 }
}

# Fallback 2: first *.crt except ca.crt
if (-not $leafCrt) {
  $anyCrt = Get-ChildItem -Path $certDir -Filter *.crt -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "ca.crt" } | Select-Object -First 1
  if ($anyCrt) { $leafCrt = $anyCrt.FullName }
}

# Ensure we have a leaf cert
if (-not $leafCrt) {
  $files = (Get-ChildItem $certDir -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }) -join ", "
  Err ("Missing leaf certificate (*.crt). Files in " + $certDir + ": " + $files)
  exit 1
}

# 4) Import root CA if present
if (Test-Path $caCrt) {
  Info "Importing root CA into Local Machine -> Trusted Root Certification Authorities..."
  & certutil -addstore -f "Root" $caCrt | Out-Null
  if ($LASTEXITCODE -eq 0) { Ok "Root CA imported." } else { Warn "Root CA import returned non-zero (may already exist)." }
} else {
  Warn "Skipped root CA import (no ca.crt). If trust still fails, upgrade office-addin-dev-certs."
}

# 5) Import leaf cert into CurrentUser stores
Info ("Importing leaf cert (" + [System.IO.Path]::GetFileName($leafCrt) + ") into CurrentUser -> My...")
& certutil -addstore -user "My" $leafCrt | Out-Null
if ($LASTEXITCODE -eq 0) { Ok "Installed into CurrentUser\My." } else { Warn "Non-zero return (may already exist)." }

Info "Importing leaf cert into CurrentUser -> TrustedPublisher..."
& certutil -addstore -user "TrustedPublisher" $leafCrt | Out-Null
if ($LASTEXITCODE -eq 0) { Ok "Installed into CurrentUser\TrustedPublisher." } else { Warn "Non-zero return (may already exist)." }

# 6) Allow WebView/Office loopback
Info "Allowing loopback for Office Desktop and Win32 WebViewHost..."
& CheckNetIsolation LoopbackExempt -a -n="Microsoft.Office.Desktop_8wekyb3d8bbwe" | Out-Null
& CheckNetIsolation LoopbackExempt -a -n="Microsoft.Win32WebViewHost_cw5n1h2txyewy" | Out-Null
Ok "Loopback exemptions added."

# 7) Basic verification
Info "Verifying stores (names may differ by version):"
try { & certutil -store root "Developer CA" | Out-Null; Ok "Root store query completed."; } catch { Warn "Root store query did not find Developer CA (can be normal without ca.crt)." }
try { & certutil -store TrustedPublisher localhost | Out-Null; Ok "TrustedPublisher query for 'localhost' completed (if you used 127.0.0.1 it may be empty)." } catch { Warn "TrustedPublisher query did not find 'localhost' (may be 127.0.0.1 or another leaf name)." }

Info "Opening https://localhost:3000 in Edge for browser-level trust check..."
Start-Process msedge "https://localhost:3000" -ErrorAction SilentlyContinue

Ok "Setup finished."

# 8) Optional auto-start
if ($StartDesktop -and $StartWeb) {
  Warn "Both -StartDesktop and -StartWeb were provided; starting desktop only."
}

if ($StartDesktop) {
  Info "Starting: npm start (desktop sideload)..."
  Start-Process powershell -ArgumentList "-NoProfile -Command npm start"
} elseif ($StartWeb) {
  Info "Starting: npm run start:web (web sideload)..."
  Start-Process powershell -ArgumentList "-NoProfile -Command npm run start:web"
} else {
  Info "To start debugging now, run: npm start    or    npm run start:web"
}
