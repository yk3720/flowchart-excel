# -*- coding: utf-8 -*-
"""PyInstaller ビルド — Flowchart Excel."""

import os
import subprocess
import sys
from pathlib import Path

script_dir = Path(__file__).resolve().parent
os.chdir(script_dir)

pyinstaller_cmd = [
    "pyinstaller",
    "--onefile",
    "--windowed",
    "--name=FlowchartExcel",
    "--icon=NONE",
    "main.py",
]

print("=" * 60)
print("Flowchart Excel — PyInstaller build")
print("=" * 60)
print(f"cwd: {script_dir}")
print(f"cmd: {' '.join(pyinstaller_cmd)}")
print("=" * 60)

result = subprocess.run(pyinstaller_cmd, check=False)
sys.exit(result.returncode)
