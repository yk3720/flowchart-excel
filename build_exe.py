# -*- coding: utf-8 -*-
"""PyInstaller ビルド — Flowchart Excel.

MUST: このスクリプトを実行した Python（sys.executable）で PyInstaller を動かす。
PATH の素の `pyinstaller` を呼ぶと、ストア版等の別環境で解析され
pywebview が同梱されずプレビューが開かなくなる（PYTHON_RULES §13）。
"""

import os
import subprocess
import sys
from pathlib import Path

script_dir = Path(__file__).resolve().parent
os.chdir(script_dir)

# ビルド interpreter に webview が入っているか先に確認（PATH 取り違えの早期検知）
try:
    import webview  # noqa: F401
except ImportError:
    print(
        "ERROR: webview (pywebview) is not importable in this Python.\n"
        f"  executable: {sys.executable}\n"
        "  Fix: .\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt\n"
        "       then: .\\.venv\\Scripts\\python.exe build_exe.py",
        file=sys.stderr,
    )
    sys.exit(1)

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

# delayed/optional import の webview は Analysis が見落とすため hidden-import 必須
pyinstaller_cmd = [
    sys.executable,
    "-m",
    "PyInstaller",
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name=FlowchartExcel",
    "--collect-all=customtkinter",
    "--collect-all=webview",
    "--hidden-import=webview",
    "--hidden-import=pythonnet",
    "--hidden-import=clr_loader",
    f"--add-data={add_preview}",
    "main.py",
]

print("=" * 60)
print("Flowchart Excel - PyInstaller build")
print("=" * 60)
print(f"cwd: {script_dir}")
print(f"python: {sys.executable}")
print(f"cmd: {' '.join(pyinstaller_cmd)}")
print("=" * 60)

result = subprocess.run(pyinstaller_cmd, check=False)
if result.returncode != 0:
    sys.exit(result.returncode)

warn = script_dir / "build" / "FlowchartExcel" / "warn-FlowchartExcel.txt"
if warn.is_file():
    text = warn.read_text(encoding="utf-8", errors="replace")
    if "missing module named webview" in text:
        print(
            "ERROR: warn-FlowchartExcel.txt still reports missing webview.\n"
            "  Do not distribute this exe — preview will fail with pywebview missing.",
            file=sys.stderr,
        )
        sys.exit(1)

print("Build OK. Smoke-check preview host before distributing (see README).")
sys.exit(0)
