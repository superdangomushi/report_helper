#!/usr/bin/env bash
# install.sh — set up report_helper on macOS.
# Prerequisite: Homebrew (https://brew.sh) is already installed.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

log()  { printf "==> %s\n" "$*"; }
fail() { printf "Error: %s\n" "$*" >&2; exit 1; }

# 1. Verify Homebrew is available.
if ! command -v brew >/dev/null 2>&1; then
  fail "Homebrew が見つかりません。https://brew.sh の手順でインストールしてから再実行してください。"
fi
log "Homebrew: $(brew --version | head -1)"

# 2. Install Node.js if missing.
if ! command -v node >/dev/null 2>&1; then
  log "Node.js が無いので brew install node を実行します。"
  brew install node
else
  log "Node.js: $(node --version)"
fi

# 3. Install npm dependencies.
log "npm install を実行します。"
npm install

# 4. Link 'mkreport' globally.
log "npm link で mkreport コマンドを使えるようにします。"
npm link

# 5. Smoke test.
if ! command -v mkreport >/dev/null 2>&1; then
  fail "mkreport コマンドが PATH に通っていません。シェルを再起動して再確認してください。"
fi

log "完了しました。"
echo ""
echo "使い方:"
echo "  mkreport start my_first_report   # 新規プロジェクトを作る"
echo "  cd my_first_report"
echo "  # README.md を参考に各ファイルを編集"
echo "  mkreport .                       # レポートを生成"
