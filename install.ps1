# install.ps1 — set up report_helper on Windows.
# install.bat をダブルクリックすれば全自動で:
#   1. 管理者権限への自己昇格
#   2. Node.js (LTS) を winget で自動インストール
#   3. PATH をその場で再読み込み
#   4. npm install / npm link
#   5. mkreport コマンドが使えるか自己診断
# だけで完了します。事前準備は不要です。

$ErrorActionPreference = "Stop"

function Log($msg)  { Write-Host "==> $msg" -ForegroundColor Cyan }
function Warn($msg) { Write-Host "!!  $msg" -ForegroundColor Yellow }
function Fail($msg) {
  Write-Host "Error: $msg" -ForegroundColor Red
  Read-Host "Enter キーで終了します"
  exit 1
}

function Refresh-Path {
  # winget や Node のインストール後、現在のシェルへ PATH を即時反映する
  $machine = [System.Environment]::GetEnvironmentVariable("Path","Machine")
  $user    = [System.Environment]::GetEnvironmentVariable("Path","User")
  $env:Path = "$machine;$user"
}

# --- 1. 管理者権限へ自己昇格 ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Log "管理者権限で実行し直します。UAC のダイアログが出たら「はい」を押してください。"
  $argsList = @("-NoProfile","-ExecutionPolicy","Bypass","-File",$PSCommandPath)
  Start-Process -FilePath "powershell.exe" -ArgumentList $argsList -Verb RunAs
  exit 0
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir
Log "作業ディレクトリ: $ScriptDir"

# --- 2. Node.js を確認、無ければ winget で自動インストール ---
Refresh-Path
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Fail "winget が見つかりません。Windows 10 1809 以降 / Windows 11 が必要です。Microsoft Store で『アプリ インストーラー』を更新してください。"
  }
  Log "Node.js (LTS) を winget でインストールします。"
  $wingetArgs = @(
    "install","--id","OpenJS.NodeJS.LTS","-e","--silent",
    "--accept-package-agreements","--accept-source-agreements"
  )
  & winget @wingetArgs
  if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne -1978335189) {
    Fail "winget による Node.js のインストールに失敗しました (exit code $LASTEXITCODE)。"
  }
  Refresh-Path
  if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    # winget の標準インストール先を PATH に追加して再試行
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
    Fail "Node.js のインストールは完了しましたが、PATH に反映されません。PC を再起動してから install.bat を再実行してください。"
  }
}
Log "Node.js: $(node --version)"
Log "npm:     $(npm --version)"

# --- 3. 依存パッケージのインストール ---
Log "npm install を実行します。"
& npm install
if ($LASTEXITCODE -ne 0) { Fail "npm install に失敗しました (exit code $LASTEXITCODE)。" }

# --- 4. mkreport コマンドをグローバルに登録 ---
Log "npm link で mkreport コマンドを登録します。"
& npm link
if ($LASTEXITCODE -ne 0) { Fail "npm link に失敗しました (exit code $LASTEXITCODE)。" }
Refresh-Path

# --- 5. 自己診断 ---
if (-not (Get-Command mkreport -ErrorAction SilentlyContinue)) {
  # npm の global bin を明示的に PATH に足してリトライ
  try {
    $npmPrefix = (& npm config get prefix).Trim()
    if ($npmPrefix -and (Test-Path $npmPrefix)) {
      $env:Path = "$npmPrefix;$env:Path"
    }
  } catch {}
}
if (-not (Get-Command mkreport -ErrorAction SilentlyContinue)) {
  Fail "mkreport コマンドが PATH に通っていません。一度ログオフ→ログオンしてから新しい PowerShell で 'mkreport help' を実行して確認してください。"
}

Log "完了しました。"
Write-Host ""
Write-Host "使い方:" -ForegroundColor Green
Write-Host "  mkreport start my_first_report   # 新規プロジェクトを作る"
Write-Host "  cd my_first_report"
Write-Host "  # README.md を参考にファイルを編集"
Write-Host "  mkreport .                       # レポートを生成"
Write-Host ""
Read-Host "Enter キーで終了します"
