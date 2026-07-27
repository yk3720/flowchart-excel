"""WebView2 プレビューホスト（--flowchart-preview で起動）。"""
from __future__ import annotations

import json
import logging
import sys
import threading
import time
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

    from app.core.live_preview import (
        LIVE_POLL_INTERVAL_SEC,
        table_fingerprint,
        try_refresh_studio_payload,
    )
    from app.core.preview_inject import build_payload_inject_js

    payload: dict[str, Any] = json.loads(payload_path.read_text(encoding="utf-8"))
    index = (dist_dir / "index.html").resolve().as_uri()
    stop_live = threading.Event()
    state = {"payload": payload, "fp": table_fingerprint(payload)}

    class Api:
        def confirm(self) -> None:
            stop_live.set()
            # P2: 確定直前に再読込してから凍結（表示＝作成）
            fresh = try_refresh_studio_payload(state["payload"])
            if fresh:
                state["payload"] = fresh
                state["fp"] = table_fingerprint(fresh)
            result_path.write_text(
                json.dumps(
                    {
                        "action": "confirm",
                        "fingerprint": state["fp"],
                        "payload": state["payload"],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            for window in webview.windows:
                window.destroy()

        def cancel(self) -> None:
            stop_live.set()
            result_path.write_text(
                json.dumps({"action": "cancel"}), encoding="utf-8"
            )
            for window in webview.windows:
                window.destroy()

    api = Api()
    # ルート C POC: Excel フォーカス時もプレビューを作業継続可能にする
    window = webview.create_window(
        title="フロープレビュー",
        url=index,
        js_api=api,
        width=1000,
        height=720,
        min_size=(720, 480),
        on_top=True,
    )

    def _inject(next_payload: dict[str, Any]) -> None:
        js = build_payload_inject_js(next_payload)
        try:
            if webview.windows:
                webview.windows[0].evaluate_js(js)
        except Exception as exc:  # noqa: BLE001 — UI 閉じ際の例外を握る
            logger.debug("inject_failed | %s", exc)

    def _on_loaded() -> None:
        _inject(state["payload"])

    def _on_closed() -> None:
        stop_live.set()
        if not result_path.exists():
            result_path.write_text(
                json.dumps({"action": "cancel"}), encoding="utf-8"
            )

    def _live_loop() -> None:
        # GUI 起動直後は注入完了を待つ
        time.sleep(1.0)
        while not stop_live.is_set() and webview.windows:
            time.sleep(LIVE_POLL_INTERVAL_SEC)
            if stop_live.is_set() or not webview.windows:
                break
            fresh = try_refresh_studio_payload(state["payload"])
            if not fresh:
                continue
            fp = table_fingerprint(fresh)
            if fp == state["fp"]:
                continue
            state["payload"] = fresh
            state["fp"] = fp
            logger.info(
                "live_preview_updated | nodes=%s",
                (fresh.get("meta") or {}).get("nodeCount"),
            )
            _inject(fresh)

    window.events.loaded += _on_loaded
    window.events.closed += _on_closed

    webview.start(func=_live_loop)
    stop_live.set()
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
