"""flowchart-excel 起動エントリポイント。"""
import os
import sys

import pythoncom
from pathlib import Path

if not getattr(sys, "frozen", False):
    if hasattr(sys.modules[__name__], "__file__") and __file__:
        sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.utils.path_utils import ensure_environment
from app.utils.logging_config import setup_logger
from app.ui.main_window import FlowchartApp
from app.constants import APP_NAME, REVISION, AUTHOR


def main() -> None:
    ensure_environment()
    logger = setup_logger()
    logger.info("starting_app | %s %s | by %s", APP_NAME, REVISION, AUTHOR)

    if os.name != "nt":
        logger.error("unsupported_os | Windows is required for Excel COM.")
        print("Error: Windows is required for Excel COM.")
        sys.exit(1)

    pythoncom.CoInitialize()
    try:
        app = FlowchartApp()
        app.mainloop()
    except KeyboardInterrupt:
        logger.info("app_interrupted_by_user")
        print("アプリケーションが中断されました。")
    except Exception as exc:
        logger.exception("unhandled_exception_at_main")
        print(f"致命的なエラーが発生しました: {exc}")
    finally:
        pythoncom.CoUninitialize()
        logger.info("app_terminated")


if __name__ == "__main__":
    main()
