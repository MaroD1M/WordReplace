#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DATA_DIR="${ROOT_DIR}/data"

if [ "$#" -lt 1 ]; then
  echo "Usage: $0 <path-to-backup.tar.gz>" >&2
  exit 1
fi

ARCHIVE="$1"

if [ ! -f "$ARCHIVE" ]; then
  echo "[restore] ERROR: archive not found: $ARCHIVE" >&2
  exit 1
fi

if ! tar -tzf "$ARCHIVE" >/dev/null 2>&1; then
  echo "[restore] ERROR: not a valid .tar.gz archive: $ARCHIVE" >&2
  exit 1
fi

cat <<EOF
[restore] WARNING: This will overwrite the contents of:
          $DATA_DIR
[restore] Stop the container before restoring:
          docker compose stop wordreplace
EOF

printf '[restore] Type "yes" to continue: '
read -r reply
if [ "$reply" != "yes" ]; then
  echo "[restore] Aborted."
  exit 1
fi

mkdir -p "$DATA_DIR"
find "$DATA_DIR" -mindepth 1 -maxdepth 1 -exec rm -rf {} +
tar -xzf "$ARCHIVE" -C "$ROOT_DIR"

echo "[restore] Restore complete from: $ARCHIVE"
echo "[restore] Start the container again with: docker compose up -d wordreplace"
