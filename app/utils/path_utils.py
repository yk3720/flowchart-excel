"""パス解決および環境診断ユーティリティ。

Powered by Auto (Cursor) (rev014)
配布（exe化）環境と開発環境の差異を吸収し、物理パスの整合性を保証する。
日本語を含むパスにも対応（UTF-8エンコーディング対応）。
"""
import sys
import os
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


def _safe_path(path_str: str) -> Path:
    """日本語を含むパスを安全に処理する。
    
    Args:
        path_str (str): パス文字列。
        
    Returns:
        Path: 安全に処理されたPathオブジェクト。
    """
    # pathlib.Pathは内部的にUTF-8をサポートしているが、
    # Windows環境での互換性を確保するため、明示的に処理
    try:
        # 文字列をそのままPathに変換（pathlibが自動的にUTF-8で処理）
        return Path(path_str)
    except (UnicodeEncodeError, UnicodeDecodeError) as e:
        logger.error(f"path_encoding_error | path={path_str[:50]} | error={e}")
        # フォールバック: 短いパス名（8.3形式）を使用
        try:
            import win32api
            short_path = win32api.GetShortPathName(path_str)
            return Path(short_path)
        except (ImportError, OSError, AttributeError) as e:
            # 最終フォールバック: 元のパスをそのまま返す
            logger.warning(f"short_path_fallback_failed | error={e}")
            return Path(path_str)


def get_app_root() -> Path:
    """アプリケーションの実行ルートディレクトリを取得する。
    
    exe化されている場合はexeの保存場所、
    スクリプト実行の場合は main.py のある場所を返す。
    日本語を含むパスにも対応。
    
    Returns:
        Path: アプリケーションルートの絶対パス。
    """
    if getattr(sys, 'frozen', False):
        # exe実行時: exeファイルのディレクトリ
        # sys.executableは文字列として取得されるため、安全に処理
        exe_path = str(sys.executable)
        return _safe_path(exe_path).parent.resolve()
    else:
        # スクリプト実行時: main.py のあるディレクトリ
        # プロジェクトルート (app/ の親) を取得
        # __file__は既にPathオブジェクトとして処理されるため、resolve()で正規化
        return Path(__file__).resolve().parent.parent.parent


def get_resource_path(relative_path: str) -> Path:
    """exe内部に埋め込まれたリソース、または外部アセットのパスを解決する。
    
    Args:
        relative_path (str): ルートからの相対パス（英語のみ推奨）。
        
    Returns:
        Path: 解決された絶対パス。
    """
    if getattr(sys, 'frozen', False):
        # PyInstallerの一時展開先 (_MEIPASS) を取得
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass:
            base_path = _safe_path(str(meipass))
        else:
            base_path = get_app_root()
        return base_path / relative_path
    
    return get_app_root() / relative_path


def ensure_environment() -> None:
    """実行に必要なディレクトリ構造を自己診断し、欠損していれば修復（生成）する。
    
    Raises:
        SystemExit: ディレクトリ作成に失敗した場合（権限不足等）。
    """
    root = get_app_root()
    required_dirs = ["logs", "output"]
    
    for dir_name in required_dirs:
        dir_path = root / dir_name
        if not dir_path.exists():
            try:
                dir_path.mkdir(parents=True, exist_ok=True)
                logger.info(f"self_healing | created missing directory: {dir_name}")
            except (OSError, PermissionError) as e:
                # 権限不足等のエラー（Layer 1.5/2.1.2 規律）
                print(f"環境エラー: ディレクトリ '{dir_name}' を作成できませんでした。権限を確認してください。({e})")
                sys.exit(1)


def get_log_file_path(filename: str = "app.log") -> Path:
    """ログファイルの保存先パスを返す。
    
    Args:
        filename (str): ログファイル名。デフォルトは"app.log"。
        
    Returns:
        Path: ログファイルの絶対パス。
    """
    return get_app_root() / "logs" / filename


def get_output_dir() -> Path:
    """出力ファイルの保存先ディレクトリを返す。
    
    Returns:
        Path: 出力ディレクトリの絶対パス。
    """
    return get_app_root() / "output"
