# install.ps1 - set up report_helper on Windows.
# Double-click install.bat and it will:
#   1. self-elevate to admin
#   2. install Node.js (LTS) via winget
#   3. refresh PATH in the current shell
#   4. run npm install / npm link
#   5. verify that mkreport is available
# No manual setup is needed.

$ErrorActionPreference = "Stop"

function Log($msg)  { Write-Host "==> $msg" -ForegroundColor Cyan }
function Warn($msg) { Write-Host "!!  $msg" -ForegroundColor Yellow }
function Fail($msg) {
  Write-Host "Error: $msg" -ForegroundColor Red
  Read-Host "Press Enter to exit"
  exit 1
}

function Refresh-Path {
  # Refresh PATH in the current shell after winget or Node installation.
  $machine = [System.Environment]::GetEnvironmentVariable("Path","Machine")
  $user    = [System.Environment]::GetEnvironmentVariable("Path","User")
  $env:Path = "$machine;$user"
}

# --- 1. Self-elevate to administrator ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Log "Restarting with administrator rights. If the UAC dialog appears, choose Yes."
  $argsList = @("-NoProfile","-ExecutionPolicy","Bypass","-File",$PSCommandPath)
  Start-Process -FilePath "powershell.exe" -ArgumentList $argsList -Verb RunAs
  exit 0
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir
Log "Working directory: $ScriptDir"

# --- 2. Check Node.js, install it with winget if needed ---
Refresh-Path
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Fail "winget was not found. Windows 10 1809 or later, or Windows 11, is required. Update App Installer from the Microsoft Store."
  }
  Log "Installing Node.js (LTS) with winget."
  $wingetArgs = @(
    "install","--id","OpenJS.NodeJS.LTS","-e","--silent",
    "--accept-package-agreements","--accept-source-agreements"
  )
  & winget @wingetArgs
  if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne -1978335189) {
    Fail "winget failed to install Node.js (exit code $LASTEXITCODE)."
  }
  Refresh-Path
  if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    # Add common winget install locations to PATH and try again.
    $candidates = @(
      "$env:ProgramFiles\nodejs",
      "${env:ProgramFiles(x86)}\nodejs",
      "$env:LOCALAPPDATA\Programs\nodejs"
    )
    foreach ($p in $candidates) {
      if (Test-Path (Join-Path $p "node.exe")) {
        $env:Path = "$p;$env:Path"
        break
      }
    }
  }
  if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Fail "Node.js installed, but it was not found on PATH. Restart the PC and run install.bat again."
  }
}
Log "Node.js: $(node --version)"
Log "npm:     $(npm --version)"

# --- 3. Install dependencies ---
Log "Running npm install."
& npm install
if ($LASTEXITCODE -ne 0) { Fail "npm install failed (exit code $LASTEXITCODE)." }

# --- 4. Register the mkreport command globally ---
Log "Registering the mkreport command with npm link."
& npm link
if ($LASTEXITCODE -ne 0) { Fail "npm link failed (exit code $LASTEXITCODE)." }
Refresh-Path

# --- 5. Self-check ---
if (-not (Get-Command mkreport -ErrorAction SilentlyContinue)) {
  # Add the npm global bin directory to PATH and try again.
  try {
    $npmPrefix = (& npm config get prefix).Trim()
    if ($npmPrefix -and (Test-Path $npmPrefix)) {
      $env:Path = "$npmPrefix;$env:Path"
    }
  } catch {}
}
if (-not (Get-Command mkreport -ErrorAction SilentlyContinue)) {
  Fail "mkreport is not on PATH. Log off and back on, then run 'mkreport help' in a new PowerShell window."
}

Log "Done."
Write-Host ""
Write-Host "Usage:" -ForegroundColor Green
Write-Host "  mkreport start my_first_report   # create a new project"
Write-Host "  cd my_first_report"
Write-Host "  # edit the files using README.md as a guide"
Write-Host "  mkreport .                       # generate the report"
Write-Host ""
Read-Host "Press Enter to exit"
