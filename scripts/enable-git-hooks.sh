#!/usr/bin/env bash
set -euo pipefail

chmod +x .githooks/pre-commit
git config core.hooksPath .githooks
echo "Git hooks enabled: core.hooksPath=.githooks"
