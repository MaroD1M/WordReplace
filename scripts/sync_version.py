#!/usr/bin/env python3
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
VERSION = (ROOT / "VERSION").read_text(encoding="utf-8").strip()

targets = {
    ROOT / "README.md": [
        ("wordreplace:v2.0.1", f"wordreplace:v{VERSION}"),
        ("例如 `v2.0.1`", f"例如 `v{VERSION}`"),
    ],
    ROOT / "GETTING_STARTED.md": [
        ("> 版本：v2.0.1  ", f"> 版本：v{VERSION}  "),
    ],
}

for path, replacements in targets.items():
    content = path.read_text(encoding="utf-8")
    for old, new in replacements:
        content = content.replace(old, new)
    path.write_text(content, encoding="utf-8")

print(f"synchronized version to v{VERSION}")
