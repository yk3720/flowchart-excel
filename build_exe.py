# -*- coding: utf-8 -*-
"""PyInstaller ビルド — Flowchart Excel."""

import os
import subprocess
import sys
from pathlib import Path

script_dir = Path(__file__).resolve().parent
os.chdir(script_dir)

preview_web = script_dir / "preview-web"
preview_dist = preview_web / "dist" / "index.html"
if preview_web.is_dir():
    print("Building preview-web (studio React Flow)...")
    npm = subprocess.run(
        ["npm", "run", "build"],
        cwd=preview_web,
        check=False,
        shell=True,
    )
    if npm.returncode != 0 or not preview_dist.is_file():
        print("ERROR: preview-web build failed", file=sys.stderr)
        sys.exit(npm.returncode or 1)

# Windows PyInstaller --add-data は src;dest
sep = ";" if os.name == "nt" else ":"
add_preview = f"preview-web/dist{sep}preview-web/dist"

pyinstaller_cmd = [
    "pyinstaller",
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name=FlowchartExcel",
    "--collect-all=customtkinter",
    "--collect-all=webview",
    f"--add-data={add_preview}",
    "main.py",
]

print("=" * 60)
print("Flowchart Excel - PyInstaller build")
print("=" * 60)
print(f"cwd: {script_dir}")
print(f"cmd: {' '.join(pyinstaller_cmd)}")
print("=" * 60)

result = subprocess.run(pyinstaller_cmd, check=False)
sys.exit(result.returncode)
