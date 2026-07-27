"""1窓プレビュー — tkwebview2 埋め込み（ルート A）。"""
from __future__ import annotations

import logging
import os
import tkinter as tk
from typing import Any, Callable, Dict, Optional

from app.core.live_preview import (
    LIVE_POLL_INTERVAL_SEC,
    table_fingerprint,
    try_refresh_studio_payload,
)
from app.core.preview_inject import build_payload_inject_js
from app.ui.studio_preview import resolve_preview_dist

logger = logging.getLogger("flowchart-excel")


def _ensure_pywebview_compat() -> None:
    """pywebview 5+/6+ 互換のプレースホルダ（将来の EdgeChrome 差分用）。"""
    try:
        from webview.platforms import edgechromium  # noqa: F401
    except ImportError:
        return


def _ensure_tkwebview2_compat() -> None:
    """tkwebview2 が pywebview 5+/6+ の EdgeChrome.webview API に追随していない修正。"""
    _ensure_pywebview_compat()
    try:
        from tkwebview2 import tkwebview2 as tkv_mod
    except ImportError:
        return
    if getattr(tkv_mod, "_yk_tkwebview2_compat", False):
        return

    from uuid import uuid4

    from webview.platforms.edgechromium import EdgeChrome
    from webview.window import Window
    from System.Windows.Forms import Control

    user32 = tkv_mod.user32
    windows = tkv_mod.windows

    def _fixed_webview2_init(
        self: Any, parent: Any, width: int, height: int, url: str = "", **kw: Any,
    ) -> None:
        tkv_mod.Frame.__init__(self, parent, width=width, height=height, **kw)
        control = Control()
        uid = "master" if len(windows) == 0 else "child_" + uuid4().hex[:8]
        window = Window(
            uid,
            str(id(self)),
            url=None,
            html=None,
            js_api=None,
            width=width,
            height=height,
            x=None,
            y=None,
            resizable=True,
            fullscreen=False,
            min_size=(200, 100),
            hidden=False,
            frameless=False,
            easy_drag=True,
            minimized=False,
            on_top=False,
            confirm_close=False,
            background_color="#FFFFFF",
            transparent=False,
            text_select=True,
            localization=None,
            zoomable=True,
            draggable=True,
            vibrancy=False,
        )
        self.window = window
        edge = EdgeChrome(control, window, None)
        self.web_view = edge
        self.control = control
        # pywebview 5+/6+: EdgeChrome.webview が WinForms WebView2 コントロール
        self.web = edge.webview
        windows.append(window)
        self.width = width
        self.height = height
        self.parent = parent
        self.chwnd = int(str(self.control.Handle))
        user32.SetParent(self.chwnd, self.winfo_id())
        user32.MoveWindow(self.chwnd, 0, 0, width, height, True)
        self.loaded = window.events.loaded
        self._WebView2__go_bind()
        if url != "":
            self.load_url(url)
        self.core = None
        self.web.CoreWebView2InitializationCompleted += self._WebView2__load_core

    tkv_mod.WebView2.__init__ = _fixed_webview2_init  # type: ignore[method-assign]
    tkv_mod._yk_tkwebview2_compat = True


def embedded_preview_available() -> bool:
    if os.name != "nt":
        return False
    try:
        _ensure_tkwebview2_compat()
        from tkwebview2.tkwebview2 import have_runtime

        return bool(have_runtime())
    except ImportError:
        return False
    except Exception:
        logger.exception("embedded_preview_runtime_check_failed")
        return False


class EmbeddedStudioPreview:
    """CTk 内に studio React Flow プレビューを埋め込む。"""

    def __init__(
        self,
        master: tk.Misc,
        schedule_after: Callable[..., str],
        on_payload_change: Optional[Callable[[], None]] = None,
    ) -> None:
        _ensure_tkwebview2_compat()
        from tkwebview2.tkwebview2 import WebView2

        self._schedule_after = schedule_after
        self._on_payload_change = on_payload_change
        self._dist = resolve_preview_dist()
        self._payload: Optional[Dict[str, Any]] = None
        self._fp = ""
        self._live = False
        self._core_ready = False
        self._page_loaded = False

        self._host = tk.Frame(master, highlightthickness=0)
        self._host.pack(fill="both", expand=True)

        self._frame = WebView2(self._host, width=800, height=480)
        self._frame.pack(fill="both", expand=True)
        self._frame.event_core_completed(self._on_core_ready)
        self._frame.loaded += self._on_page_loaded

        if self._dist is None:
            logger.warning("embedded_preview_dist_missing")
            return

        index_uri = (self._dist / "index.html").resolve().as_uri()
        self._frame.load_url(index_uri)

    @property
    def host(self) -> tk.Frame:
        return self._host

    @property
    def has_payload(self) -> bool:
        return self._payload is not None

    @property
    def payload(self) -> Optional[Dict[str, Any]]:
        return self._payload

    def load_session(self, payload: Dict[str, Any]) -> None:
        meta = dict(payload.get("meta") or {})
        meta["embedded"] = True
        meta["live"] = True
        payload = dict(payload)
        payload["meta"] = meta
        self._payload = payload
        self._fp = table_fingerprint(payload)
        self._inject(payload)
        self.start_live()
        if self._on_payload_change:
            self._on_payload_change()

    def start_live(self) -> None:
        self._live = True
        self._schedule_live()

    def stop_live(self) -> None:
        self._live = False

    def freeze_snapshot(self) -> Optional[Dict[str, Any]]:
        if not self._payload:
            return None
        fresh = try_refresh_studio_payload(self._payload)
        if fresh:
            meta = dict(fresh.get("meta") or {})
            meta["embedded"] = True
            fresh["meta"] = meta
            self._payload = fresh
            self._fp = table_fingerprint(fresh)
        return self._payload

    def is_create_enabled(self) -> bool:
        if not self._payload:
            return False
        return bool((self._payload.get("meta") or {}).get("nodeCount", 0))

    def _on_core_ready(self, sender: Any, _args: Any) -> None:
        self._core_ready = True
        if self._payload and self._page_loaded:
            self._inject(self._payload)

    def _on_page_loaded(self) -> None:
        self._page_loaded = True
        if self._payload:
            self._inject(self._payload)

    def _inject(self, payload: Dict[str, Any]) -> None:
        if not self._core_ready or not self._page_loaded:
            return
        try:
            self._frame.evaluate_js(build_payload_inject_js(payload))
        except Exception as exc:  # noqa: BLE001
            logger.debug("embedded_inject_failed | %s", exc)

    def _schedule_live(self) -> None:
        if not self._live:
            return
        self._schedule_after(int(LIVE_POLL_INTERVAL_SEC * 1000), self._live_tick)

    def _live_tick(self) -> None:
        if not self._live or not self._payload:
            return
        fresh = try_refresh_studio_payload(self._payload)
        if fresh:
            meta = dict(fresh.get("meta") or {})
            meta["embedded"] = True
            meta["live"] = True
            fresh["meta"] = meta
            fp = table_fingerprint(fresh)
            if fp != self._fp:
                self._payload = fresh
                self._fp = fp
                logger.info(
                    "embedded_live_updated | nodes=%s",
                    (fresh.get("meta") or {}).get("nodeCount"),
                )
                self._inject(fresh)
                if self._on_payload_change:
                    self._on_payload_change()
        self._schedule_live()
