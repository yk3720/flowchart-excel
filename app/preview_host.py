"""WebView2 プレビューホスト（--flowchart-preview で起動）。"""
from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

logger = logging.getLogger("flowchart-excel-preview-host")


def run_preview_host(payload_path: Path, result_path: Path, dist_dir: Path) -> int:
    try:
        import webview
    except ImportError:
        logger.error("pywebview_missing")
        result_path.write_text(
            json.dumps({"action": "cancel", "error": "pywebview missing"}),
            encoding="utf-8",
        )
        return 2

    payload: dict[str, Any] = json.loads(payload_path.read_text(encoding="utf-8"))
    index = (dist_dir / "index.html").resolve().as_uri()

    class Api:
        def confirm(self) -> None:
            result_path.write_text(
                json.dumps({"action": "confirm"}), encoding="utf-8"
            )
            for window in webview.windows:
                window.destroy()

        def cancel(self) -> None:
            result_path.write_text(
                json.dumps({"action": "cancel"}), encoding="utf-8"
            )
            for window in webview.windows:
                window.destroy()

    api = Api()
    window = webview.create_window(
        title="フロープレビュー — 確認してから作成",
        url=index,
        js_api=api,
        width=1000,
        height=720,
        min_size=(720, 480),
    )

    payload_json = json.dumps(payload, ensure_ascii=False)

    def _on_loaded() -> None:
        # React mount 前後どちらでも届くようリトライ注入
        js = f"""
        window.__PREVIEW_PAYLOAD__ = {payload_json};
        (function tryInject(n) {{
          if (window.setPreviewPayload) {{
            window.setPreviewPayload(window.__PREVIEW_PAYLOAD__);
          }} else if (n < 100) {{
            setTimeout(function() {{ tryInject(n + 1); }}, 50);
          }}
        }})(0);
        """
        window.evaluate_js(js)

    def _on_closed() -> None:
        if not result_path.exists():
            result_path.write_text(
                json.dumps({"action": "cancel"}), encoding="utf-8"
            )

    window.events.loaded += _on_loaded
    window.events.closed += _on_closed

    webview.start()
    return 0


def main(argv: list[str] | None = None) -> int:
    args = list(argv if argv is not None else sys.argv[1:])
    if len(args) < 3:
        print(
            "Usage: --flowchart-preview <payload.json> <result.json> <dist_dir>",
            file=sys.stderr,
        )
        return 2
    return run_preview_host(Path(args[0]), Path(args[1]), Path(args[2]))


if __name__ == "__main__":
    raise SystemExit(main())
