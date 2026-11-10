#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DATA_DIR="${ROOT_DIR}/data"
BACKUP_DIR="${ROOT_DIR}/backups"

if [ ! -d "$DATA_DIR" ]; then
  echo "[backup] ERROR: data directory not found: $DATA_DIR" >&2
  echo "[backup] Start the container once so ./data is created, then retry." >&2
  exit 1
fi

mkdir -p "$BACKUP_DIR"

STAMP="$(date +%Y%m%d-%H%M%S)"
ARCHIVE="${BACKUP_DIR}/wordreplace-data-${STAMP}.tar.gz"
TMP_ARCHIVE="${ARCHIVE}.tmp"

trap 'rm -f "$TMP_ARCHIVE"' EXIT

tar -czf "$TMP_ARCHIVE" -C "$ROOT_DIR" data
mv "$TMP_ARCHIVE" "$ARCHIVE"
trap - EXIT

echo "[backup] Backup created: $ARCHIVE"
if command -v du >/dev/null 2>&1; then
  echo "[backup] Size: $(du -h "$ARCHIVE" | cut -f1)"
fi
echo "[backup] Restore with: ./scripts/restore.sh \"$ARCHIVE\""
