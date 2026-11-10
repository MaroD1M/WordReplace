#!/usr/bin/env bash
set -uo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

declare -a RESULTS=()
FAILED=0

record() {
  local status="$1"
  local name="$2"
  RESULTS+=("${status}  ${name}")
  if [ "$status" = "FAIL" ]; then
    FAILED=1
  fi
}

run_check() {
  local name="$1"
  shift
  printf '\n==> %s\n' "$name"
  if "$@"; then
    record PASS "$name"
  else
    record FAIL "$name"
  fi
}

run_check "Version sync" python3 scripts/sync_version.py --check

if [ -d backend/.venv ]; then
  run_check "Backend tests (pytest)" bash -c 'cd backend && .venv/bin/python -m pytest -q'
else
  RESULTS+=("SKIP  Backend tests (backend/.venv not found)")
fi

if [ -d frontend/node_modules ]; then
  if command -v pnpm >/dev/null 2>&1; then
    run_check "Frontend lint/build" bash -c 'cd frontend && pnpm lint && pnpm build'
  elif command -v corepack >/dev/null 2>&1; then
    run_check "Frontend lint/build" bash -c 'cd frontend && corepack pnpm lint && corepack pnpm build'
  else
    RESULTS+=("FAIL  Frontend lint/build (pnpm not found)")
    FAILED=1
  fi
else
  RESULTS+=("SKIP  Frontend lint/build (frontend/node_modules not found)")
fi

echo
echo "==================== Summary ===================="
for result in "${RESULTS[@]}"; do
  echo "$result"
done
echo "================================================="

if [ "$FAILED" -ne 0 ]; then
  echo "Result: FAILED"
  exit 1
fi

echo "Result: PASSED"
exit 0
