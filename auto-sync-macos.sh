#!/usr/bin/env zsh
set -u

REPO_DIR="/Users/michalsip/repo/code"
BRANCH="main"
LOG_DIR="$HOME/Library/Logs/code-auto-sync"
LOCK_DIR="/tmp/code-auto-sync.lock"

mkdir -p "$LOG_DIR"

if ! mkdir "$LOCK_DIR" 2>/dev/null; then
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] Auto-sync already running" >> "$LOG_DIR/auto-sync.log"
  exit 0
fi

cleanup() {
  rmdir "$LOCK_DIR" 2>/dev/null || true
}
trap cleanup EXIT

cd "$REPO_DIR" || exit 1

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] Not a git repository: $REPO_DIR" >> "$LOG_DIR/auto-sync.log"
  exit 1
fi

git fetch origin >> "$LOG_DIR/auto-sync.log" 2>&1

if ! git diff --quiet || ! git diff --cached --quiet || [ -n "$(git ls-files --others --exclude-standard)" ]; then
  git add -A
  git commit -m "auto-sync: MacBook $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG_DIR/auto-sync.log" 2>&1
fi

git pull --rebase origin "$BRANCH" >> "$LOG_DIR/auto-sync.log" 2>&1
git push origin "$BRANCH" >> "$LOG_DIR/auto-sync.log" 2>&1

echo "[$(date '+%Y-%m-%d %H:%M:%S')] Auto-sync finished" >> "$LOG_DIR/auto-sync.log"