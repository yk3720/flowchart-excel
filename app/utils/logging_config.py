"""構造化ロギング設定。

Powered by Auto (Cursor) (rev014)
"""
import logging
import sys
from logging.handlers import RotatingFileHandler
from datetime import datetime
from app.utils.path_utils import get_log_file_path


class StructuredFormatter(logging.Formatter):
    """Key-Value形式の構造化ロガー。"""
    
    def format(self, record: logging.LogRecord) -> str:
        """ログレコードを構造化形式にフォーマットする。
        
        Args:
            record (logging.LogRecord): ログレコード。
            
        Returns:
            str: フォーマットされたログメッセージ。
        """
        timestamp = datetime.fromtimestamp(record.created).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        level = record.levelname
        message = record.getMessage()
        
        # 構造化メッセージの組み立て
        structured_msg = f"time={timestamp} | level={level} | msg={message}"
        
        if record.exc_info:
            structured_msg += f" | exception={self.formatException(record.exc_info)}"
            
        return structured_msg


def setup_logger(name: str = "flowchart-excel") -> logging.Logger:
    """ロガーをセットアップし、コンソールとファイルの両方に出力する。
    
    Args:
        name (str): ロガー名。デフォルトは"MZ0000"。
        
    Returns:
        logging.Logger: セットアップされたロガー。
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # 既存のハンドラをクリア
    if logger.hasHandlers():
        logger.handlers.clear()
    
    formatter = StructuredFormatter()
    
    # 1. コンソールハンドラ
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # 2. ファイルハンドラ (RotatingFileHandler を使用 - rev014 循環型ロギング)
    try:
        log_file = get_log_file_path()
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5,          # 5世代管理
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except (OSError, PermissionError, IOError) as e:
        print(f"Warning: Could not setup file logging: {e}")
        
    return logger
