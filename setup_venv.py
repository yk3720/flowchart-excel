# -*- coding: utf-8 -*-
"""
仮想環境セットアップスクリプト
MZ0000_FlowchartTool rev014

プログラム名: MZ0000_FlowchartTool
操作AI名: Auto (Cursor)
"""

import subprocess
import sys
from pathlib import Path

# プロジェクトルート
project_root = Path(__file__).resolve().parent

# 仮想環境のパス
venv_path = project_root / ".venv"

# 現在のPythonバージョンを取得
current_python_version = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
print(f"現在のPythonバージョン: {current_python_version} ({sys.version.split()[0]})")

# .python-versionからバージョンを読み取る
python_version_file = project_root / ".python-version"
python_version = None
if python_version_file.exists():
    python_version = python_version_file.read_text(encoding="utf-8").strip()
    print(f"✓ .python-versionを検出: {python_version}")
    
    # バージョンの不一致をチェック
    if python_version != current_python_version:
        print(f"【警告】.python-version ({python_version}) と現在のPython ({current_python_version}) が異なります")
        print(f"        現在のPython ({current_python_version}) で仮想環境を作成します")
        print(f"        .python-versionを更新する場合は、手動で編集してください")

print("=" * 60)
print("仮想環境の作成を開始します")
print("=" * 60)
print(f"プロジェクトルート: {project_root}")
print(f"仮想環境パス: {venv_path}")
print(f"使用するPythonバージョン: {current_python_version}")
if python_version and python_version != current_python_version:
    print(f"（.python-version指定: {python_version} - 無視されます）")
print("=" * 60)

# 既に存在する場合は確認
if venv_path.exists():
    print(f"【警告】仮想環境は既に存在します: {venv_path}")
    response = input("上書きしますか？ (y/N): ")
    if response.lower() != 'y':
        print("処理を中止しました")
        sys.exit(0)
    import shutil
    shutil.rmtree(venv_path)
    print(f"既存の仮想環境を削除しました: {venv_path}")

# 仮想環境の作成
# 現在のPython（sys.executable）を使用して仮想環境を作成
# これにより、実行時のPythonバージョンで仮想環境が作成される
python_cmd = [sys.executable, "-m", "venv", str(venv_path)]
print(f"使用するPython: {sys.executable}")
print(f"Pythonバージョン: {current_python_version}")

try:
    subprocess.run(
        python_cmd,
        check=True,
        cwd=project_root
    )
    print(f"✓ 仮想環境を作成しました: {venv_path}")
except subprocess.CalledProcessError as e:
    print(f"【エラー】仮想環境の作成に失敗しました: {e}")
    if sys.platform == "win32" and python_version:
        print(f"【ヒント】Python {python_version} がインストールされているか確認してください")
        print(f"        インストール済みバージョンを確認: py --list")
    sys.exit(1)
except FileNotFoundError:
    print("【エラー】Pythonが見つかりません")
    if sys.platform == "win32":
        print("【ヒント】Python Launcher（py）がインストールされているか確認してください")
    sys.exit(1)

# 仮想環境のPythonパス
if sys.platform == "win32":
    venv_python = venv_path / "Scripts" / "python.exe"
else:
    venv_python = venv_path / "bin" / "python"

if not venv_python.exists():
    print(f"【エラー】仮想環境のPythonが見つかりません: {venv_python}")
    sys.exit(1)

print(f"✓ 仮想環境のPython: {venv_python}")

# pipのアップグレード
print("\n" + "=" * 60)
print("pipをアップグレードします")
print("=" * 60)
try:
    subprocess.run(
        [str(venv_python), "-m", "pip", "install", "--upgrade", "pip", "--quiet"],
        check=True,
        cwd=project_root
    )
    print("✓ pipのアップグレードが完了しました")
except subprocess.CalledProcessError as e:
    print(f"【警告】pipのアップグレードに失敗しました: {e}")
    print("続行します...")

# requirements.txtのインストール
requirements_file = project_root / "requirements.txt"
if requirements_file.exists():
    print("\n" + "=" * 60)
    print("依存パッケージをインストールします")
    print("=" * 60)
    try:
        subprocess.run(
            [str(venv_python), "-m", "pip", "install", "-r", str(requirements_file)],
            check=True,
            cwd=project_root
        )
        print("✓ 依存パッケージのインストールが完了しました")
    except subprocess.CalledProcessError as e:
        print(f"【エラー】依存パッケージのインストールに失敗しました: {e}")
        sys.exit(1)
else:
    print(f"【警告】requirements.txtが見つかりません: {requirements_file}")

print("\n" + "=" * 60)
print("仮想環境のセットアップが完了しました")
print("=" * 60)
print(f"仮想環境のパス: {venv_path}")
print(f"Pythonパス: {venv_python}")
print("\n仮想環境を有効化するには:")
if sys.platform == "win32":
    print(f"  {venv_path}\\Scripts\\activate")
else:
    print(f"  source {venv_path}/bin/activate")
print("=" * 60)
