# -*- coding: utf-8 -*-
"""
環境浄化スクリプト
MZ0000_FlowchartTool rev014

プログラム名: MZ0000_FlowchartTool
操作AI名: Auto (Cursor)
"""

import shutil
from pathlib import Path

# プロジェクトルート
project_root = Path(__file__).resolve().parent

# 削除対象のディレクトリとファイル
cleanup_targets = [
    # PyInstaller関連
    project_root / "build",
    project_root / "dist",
    project_root / "MZ0000_FlowchartTool_rev014.spec",
    
    # Pythonキャッシュ
    project_root / "__pycache__",
    project_root / "app" / "__pycache__",
    project_root / "app" / "core" / "__pycache__",
    project_root / "app" / "ui" / "__pycache__",
    project_root / "app" / "utils" / "__pycache__",
    
    # 一時ファイル
    project_root / "temp",
    
    # ログファイル（オプション: コメントアウトで保持）
    # project_root / "logs",
]

print("=" * 60)
print("環境浄化を開始します")
print("=" * 60)

for target in cleanup_targets:
    if target.exists():
        try:
            if target.is_dir():
                shutil.rmtree(target)
                print(f"削除: {target} (ディレクトリ)")
            else:
                target.unlink()
                print(f"削除: {target} (ファイル)")
        except (OSError, PermissionError) as e:
            print(f"【警告】削除に失敗しました: {target}, {e}")

print("=" * 60)
print("環境浄化が完了しました")
print("=" * 60)
