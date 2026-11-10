#!/usr/bin/env python3
"""Keep every version literal in sync with the root VERSION file.

Usage:
    python3 scripts/sync_version.py          # write updates
    python3 scripts/sync_version.py --check  # fail if anything is out of sync
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VERSION = (ROOT / "VERSION").read_text(encoding="utf-8").strip()

# (path, compiled pattern, replacement template using {version})
TARGETS = [
    (
        ROOT / "backend/app/core/config.py",
        re.compile(r'(app_version: str = ")[^"]+(")'),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "backend/app/core/config.py",
        re.compile(r'(os\.getenv\("APP_VERSION", ")[^"]+("\))'),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "frontend/package.json",
        re.compile(r'("version": ")[^"]+(")'),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "frontend/src/lib/constants.ts",
        re.compile(r'(NEXT_PUBLIC_APP_VERSION \?\? ")[^"]+(")'),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "Dockerfile",
        re.compile(r"(ARG APP_VERSION=)[^\s]+"),
        r"\g<1>{version}",
    ),
    (
        ROOT / "README.md",
        re.compile(r"(wordreplace:)(?:v)?\d+\.\d+\.\d+"),
        r"\g<1>{version}",
    ),
    (
        ROOT / "README.md",
        re.compile(r"(当前版本：\*\*v)\d+\.\d+\.\d+(\*\*)"),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "README.md",
        re.compile(r"(如 `)\d+\.\d+\.\d+(`；GitHub 发布标签为)"),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "README.md",
        re.compile(r"(例如 `)\d+\.\d+\.\d+(`)"),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "SECURITY.md",
        re.compile(r"(wordreplace:)(?:v)?\d+\.\d+\.\d+"),
        r"\g<1>{version}",
    ),
    (
        ROOT / "docs/DEPLOYMENT.md",
        re.compile(r"(wordreplace:)(?:v)?\d+\.\d+\.\d+"),
        r"\g<1>{version}",
    ),
    (
        ROOT / "docs/DEPLOYMENT.md",
        re.compile(r"(本示例为 `)\d+\.\d+\.\d+(`)"),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "docs/TROUBLESHOOTING.md",
        re.compile(r"(例如 `)\d+\.\d+\.\d+(`（GitHub)"),
        r"\g<1>{version}\g<2>",
    ),
    (
        ROOT / "GETTING_STARTED.md",
        re.compile(r"(当前版本为 \*\*v)\d+\.\d+\.\d+(\*\*)"),
        r"\g<1>{version}\g<2>",
    ),
]


def main() -> int:
    check_only = "--check" in sys.argv
    problems: list[str] = []

    for path, pattern, template in TARGETS:
        content = path.read_text(encoding="utf-8")
        replacement = template.format(version=VERSION)
        updated, count = pattern.subn(replacement, content)
        if count == 0:
            problems.append(f"{path.relative_to(ROOT)}: version pattern not found")
            continue
        if updated != content:
            if check_only:
                problems.append(f"{path.relative_to(ROOT)}: version is out of sync")
            else:
                path.write_text(updated, encoding="utf-8")

    if problems:
        print("\n".join(problems))
        return 1

    print(f"version is in sync: v{VERSION}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
