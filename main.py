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
from app.constants import APP_NAME, REVISION, AUTHOR


def _run_gui(logger) -> None:
    from app.ui.embedded_preview import _ensure_tkwebview2_compat

    _ensure_tkwebview2_compat()
    from app.ui.main_window import FlowchartApp

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


def main() -> None:
    # WebView プレビュー子プロセス（exe 同梱時も同じバイナリ · フォールバック用）
    if len(sys.argv) >= 4 and sys.argv[1] == "--flowchart-preview":
        from app.preview_host import main as preview_main

        raise SystemExit(preview_main(sys.argv[2:]))

    ensure_environment()
    logger = setup_logger()
    logger.info("starting_app | %s %s | by %s", APP_NAME, REVISION, AUTHOR)

    if os.name != "nt":
        logger.error("unsupported_os | Windows is required for Excel COM.")
        print("Error: Windows is required for Excel COM.")
        sys.exit(1)

    # tkwebview2 埋め込みは STA スレッド必須
    if os.name == "nt":
        try:
            import clr

            clr.AddReference("System.Threading")
            from System.Threading import ApartmentState, Thread, ThreadStart

            def _sta_entry() -> None:
                try:
                    _run_gui(logger)
                except Exception:
                    logger.exception("sta_thread_gui_failed")
                    raise

            thread = Thread(ThreadStart(_sta_entry))
            thread.ApartmentState = ApartmentState.STA
            thread.Start()
            thread.Join()
            return
        except Exception:
            logger.exception("sta_thread_start_failed")

    _run_gui(logger)


if __name__ == "__main__":
    main()
