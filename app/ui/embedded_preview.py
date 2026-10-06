"""1窓プレビュー — tkwebview2 埋め込み（ルート A）。"""
from __future__ import annotations

import json
import logging
import os
import threading
import time
import tkinter as tk
from collections.abc import Callable
from typing import Any

import pythoncom

from app.constants import COLOR_VALIDATION, SHAPE_TYPE_VALIDATION
from app.core.autosave_check import autosave_reminder_for_watch
from app.core.id_proposal import IdProposalResult, compute_id_proposals
from app.core.id_writer import write_id_updates
from app.core.level_inference import ProposalResult, compute_level_proposals
from app.core.level_writer import write_level_updates
from app.core.live_preview import (
    LIVE_POLL_INTERVAL_SEC,
    read_watched_range,
    table_fingerprint,
    try_refresh_studio_payload,
)
from app.core.parse_table import parse_level_optional, parse_table_rows
from app.core.preview_inject import build_payload_inject_js, build_set_global_js
from app.core.proposal_decision import decide_proposal_next_action
from app.core.proposal_fingerprint import LEVEL_COL, ProposalSnapshot, capture_snapshot
from app.core.row_validation import apply_new_row_validation
from app.ui.studio_preview import resolve_preview_dist

logger = logging.getLogger("flowchart-excel")

# 注入リトライ（setPreviewPayload 準備待ち · evaluate 失敗時）
_INJECT_RETRY_MS = 200
_INJECT_RETRY_MAX = 25

# C-2: 区別①（トポロジー変更）検知による自動再計算の上限回数
_PROPOSAL_STALE_RETRY_MAX = 3

# F8: JS→Python応答の無応答タイムアウト判定基準（秒）。WebView2のJS実行レイテンシの
# 実測値を見て調整する想定（構想設計§4）。
_JS_RESPONSE_TIMEOUT_SEC = 5.0


def _ensure_pywebview_compat() -> None:
    """pywebview 5+/6+ 互換のプレースホルダ（将来の EdgeChrome 差分用）。"""
    try:
        from webview.platforms import edgechromium  # noqa: F401
    except ImportError:
        return


def _ensure_tkwebview2_compat() -> None:
    """tkwebview2 が pywebview 5+/6+ の EdgeChrome API に追随していない修正。"""
    _ensure_pywebview_compat()
    try:
        from tkwebview2 import tkwebview2 as tkv_mod
    except ImportError:
        return
    if getattr(tkv_mod, "_yk_tkwebview2_compat", False):
        return

    from uuid import uuid4

    from System.Windows.Forms import Control
    from webview.platforms.edgechromium import EdgeChrome
    from webview.window import Window

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

    def _fixed_evaluate_js(self: Any, script: str, callback: Any = None) -> Any:
        """pywebview 5+: EdgeChrome.evaluate_js(script, parse_json) に合わせる。

        旧 tkwebview2 は (script, semaphore, js_r) を渡しており TypeError で注入失敗する。
        """
        result = self.web_view.evaluate_js(script, False)
        if callback is not None:
            callback(result)
        return result

    tkv_mod.WebView2.__init__ = _fixed_webview2_init  # type: ignore[method-assign]
    tkv_mod.WebView2.evaluate_js = _fixed_evaluate_js  # type: ignore[method-assign]
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
        on_payload_change: Callable[[], None] | None = None,
        on_js_timeout: Callable[[], None] | None = None,
        on_row_validation_applied: Callable[[int], None] | None = None,
        on_autosave_reminder: Callable[[str], None] | None = None,
    ) -> None:
        _ensure_tkwebview2_compat()
        from tkwebview2.tkwebview2 import WebView2

        self._schedule_after = schedule_after
        self._on_payload_change = on_payload_change
        self._on_js_timeout = on_js_timeout
        self._on_row_validation_applied = on_row_validation_applied
        self._on_autosave_reminder = on_autosave_reminder
        self._dist = resolve_preview_dist()
        self._payload: dict[str, Any] | None = None
        self._fp = ""
        self._live = False
        self._core_ready = False
        self._page_loaded = False
        self._inject_retries = 0
        self._pending_main_work = False
        self._pump_alive = True
        # React側 generated.ok の反映（ExecuteScriptAsync ポーリング）。未取得のうちは作成不可扱い（安全側デフォルト）
        self._validation_ok: bool | None = None
        self._pending_validation_result: tuple[bool, Any] | None = None

        # F8: 「最後に応答を受信した時刻」基準の無応答検出（構想設計§4）。
        # 「最後に送信した時刻」を基準にしないのは、`_live_tick`が無条件・定期的に
        # 送信し続けるため、WebView2が完全にハングしていても送信時刻は毎ティック
        # 更新され、閾値を超えることが永久に無いという欠陥を避けるため。
        self._last_js_response_at = time.monotonic()
        self._js_timeout_notified = False

        # C-2: 「提案」タブの JS↔Python ポーリングブリッジ（window.__proposalAction 経由）
        self._proposal_stop_event = threading.Event()
        self._pending_proposal_action: dict[str, Any] | None = None
        self._proposal_last_request_id: str | None = None
        self._proposal_baseline: ProposalSnapshot | None = None
        self._proposal_result: ProposalResult | None = None
        self._proposal_stale_retries = 0
        # F8: compute/update の応答をまだ送っていない request_id（無応答タイムアウト時に
        # ベストエフォートでUI復旧指示を送る対象の特定に使う）。
        self._proposal_inflight_request_id: str | None = None

        # F4: 新規行へのデータ入力規則継承の基準行数（構想設計§5）。Noneは未計測。
        self._f4_row_count: int | None = None

        # F5: 「ID採番」タブのJS↔Pythonポーリングブリッジ（scope="id"）。
        # `_pending_proposal_action`/`_proposal_last_request_id`/`_proposal_stop_event`は
        # scope（"level"/"id"）に関係なく共有する（window.__proposalAction自体は1つ）。
        self._id_proposal_baseline: ProposalSnapshot | None = None
        self._id_proposal_result: IdProposalResult | None = None
        self._id_proposal_row_count: int | None = None
        self._id_stale_retries = 0

        self._host = tk.Frame(master, highlightthickness=0)
        self._host.pack(fill="both", expand=True)
        # 初期 HWND は親より小さく（airspace 防止）。
        # __init__ 中の update*/winfo 禁止。CLR/loaded コールバックからも Tk を触らない。
        self._frame = WebView2(self._host, width=320, height=480)
        self._frame.pack(fill="both", expand=True)
        self._frame.event_core_completed(self._on_core_ready_clr)
        self._frame.loaded += self._on_page_loaded_bg
        self._schedule_after(50, self._main_pump)

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
    def payload(self) -> dict[str, Any] | None:
        return self._payload

    def load_session(self, payload: dict[str, Any]) -> None:
        meta = dict(payload.get("meta") or {})
        meta["embedded"] = True
        meta["live"] = True
        payload = dict(payload)
        payload["meta"] = meta
        self._payload = payload
        self._fp = table_fingerprint(payload)
        self._inject_retries = 0
        # 新規セッションごとにReact側からの再報告を待つ（安全側デフォルトへリセット）
        self._validation_ok = None
        self._f4_row_count = None
        self._inject(payload)
        self.start_live()
        if self._on_payload_change:
            self._on_payload_change()

    def start_live(self) -> None:
        self._live = True
        self._schedule_live()

    def stop_live(self) -> None:
        self._live = False

    def freeze_snapshot(self) -> dict[str, Any] | None:
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

    def clear_session(self) -> None:
        """プレビューを破棄し、待機状態のWebViewへ戻す（キャンセル用）。"""
        self.stop_live()
        self._payload = None
        self._fp = ""
        self._inject_retries = 0
        self._validation_ok = None
        self._f4_row_count = None
        self._proposal_stop_event.set()
        self._proposal_baseline = None
        self._proposal_result = None
        self._id_proposal_baseline = None
        self._id_proposal_result = None
        self._id_proposal_row_count = None
        if self._core_ready:
            js = "window.setPreviewPayload && window.setPreviewPayload(null);"
            try:
                core = getattr(self._frame, "core", None)
                if core is None:
                    core = getattr(self._frame.web, "CoreWebView2", None)
                if core is not None:
                    core.ExecuteScriptAsync(js)
                else:
                    self._frame.evaluate_js(js)
            except Exception as exc:  # noqa: BLE001
                logger.debug("embedded_clear_failed | %s", exc)
        if self._on_payload_change:
            self._on_payload_change()

    def is_create_enabled(self) -> bool:
        if not self._payload:
            return False
        has_nodes = bool((self._payload.get("meta") or {}).get("nodeCount", 0))
        return has_nodes and bool(self._validation_ok)

    def _check_js_response_timeout(self) -> None:
        """構想設計§4: 無応答タイムアウトの検出とUI復旧の分離（検出部分）。

        `_poll_validation`/`_poll_proposal_action`の両方から`_live_tick`で呼ばれる
        共通の検出ロジック。「最後に応答を受信した時刻」が閾値を超えて更新されて
        いなければ無応答とみなし、`_handle_js_response_timeout`（UI復旧側）へ一度だけ
        委譲する（応答が戻るまで毎ティック繰り返し通知しない）。
        """
        timed_out = (time.monotonic() - self._last_js_response_at) > _JS_RESPONSE_TIMEOUT_SEC
        if not timed_out:
            self._js_timeout_notified = False
            return
        if self._js_timeout_notified:
            return
        self._js_timeout_notified = True
        self._handle_js_response_timeout()

    def _handle_js_response_timeout(self) -> None:
        """構想設計§4: 無応答タイムアウト検出時のUI復旧（検出と分離した応答側）。

        ステータス行の文言表示（Tkinter直接・WebViewを経由しない）は無応答時でも
        必ず届く。一方、React側のボタンを再試行可能な状態に戻す指示は
        `_push_update_result`経由（`_poll_proposal_action`と同一チャネル）のため、
        WebView2のレンダラー自体が完全にハングしている場合は届かない可能性がある
        （既知の限界）。そのため呼び出し元（`main_window.py`）には同一チャネルに
        依存しない復旧手段（ウィンドウの開き直し）を案内する文言を使わせる。
        """
        logger.warning(
            "embedded_js_response_timeout | inflight_request=%s",
            self._proposal_inflight_request_id,
        )
        if self._on_js_timeout:
            self._on_js_timeout()
        request_id = self._proposal_inflight_request_id
        if request_id is not None:
            self._push_update_result(
                request_id,
                ok=False,
                error=(
                    "プレビューが応答していません。数秒待っても改善しない場合は"
                    "プレビュー画面を閉じて開き直してください。"
                ),
            )

    def _poll_js(self, script: str, on_result: Callable[[Any], None]) -> None:
        """ExecuteScriptAsync + ContinueWith の共通送信処理（構想設計§4）。

        `_poll_validation`/`_poll_proposal_action`はcore取得→送信→コールバック登録
        という構造がほぼ同一のため、ここに共通化する。

        この埋め込みWebViewは pywebview の正規初期化（webview.start()）を経由せず
        window.gui が None のままのため、window.expose() が内部で使う window.run_js()
        は例外を出す（inject_pywebview 全般が機能しない）。また pywebview の
        evaluate_js（EdgeChrome内部でControl.Invoke+semaphoreで同期待機する実装）は
        このアプリにWinFormsのメッセージポンプが無いため呼び出したスレッドを
        永久にブロックする（実機検証で確認済み・Tkメインスレッドがフリーズする）。
        そのため _inject と同じ core.ExecuteScriptAsync を直接叩き、結果は
        ContinueWith の非同期コールバックで受け取る。
        """
        core = getattr(self._frame, "core", None)
        if core is None:
            core = getattr(self._frame.web, "CoreWebView2", None)
        if core is None:
            return
        try:
            from System import Action, String
            from System.Threading.Tasks import Task

            task = core.ExecuteScriptAsync(script)
            task.ContinueWith(Action[Task[String]](on_result))
        except Exception as exc:  # noqa: BLE001
            logger.debug("embedded_js_poll_failed | %s", exc)

    def _poll_validation(self) -> None:
        """React側の window.__flowchartValidation を読み取る（_poll_js 経由）。"""
        self._poll_js("window.__flowchartValidation || null", self._on_validation_js_result)

    def _on_validation_js_result(self, task: Any) -> None:
        """CLR/スレッドプールから呼ばれる。Tkには触れずフラグだけ立てる。"""
        try:
            raw = task.Result
        except Exception:  # noqa: BLE001
            return
        self._last_js_response_at = time.monotonic()
        try:
            data = json.loads(raw) if raw else None
        except Exception:  # noqa: BLE001
            return
        if not isinstance(data, dict):
            return
        self._pending_validation_result = (bool(data.get("ok")), data.get("errorCount"))

    def _poll_proposal_action(self) -> None:
        """C-2: React側の window.__proposalAction を読み取る（_poll_js 経由）。"""
        self._poll_js("window.__proposalAction || null", self._on_proposal_action_js_result)

    def _on_proposal_action_js_result(self, task: Any) -> None:
        """CLR/スレッドプールから呼ばれる。Tkには触れずフラグだけ立てる。"""
        try:
            raw = task.Result
        except Exception:  # noqa: BLE001
            return
        self._last_js_response_at = time.monotonic()
        try:
            data = json.loads(raw) if raw else None
        except Exception:  # noqa: BLE001
            return
        if not isinstance(data, dict):
            return
        self._pending_proposal_action = data

    def _handle_proposal_action(self, req: dict[str, Any]) -> None:
        """Tk メインスレッドから呼ばれる。実処理は background thread へ委譲する。"""
        request_id = req.get("requestId")
        scope = req.get("scope")
        action = req.get("action")
        mode = req.get("mode") or "blank_only"
        if scope not in ("level", "id"):
            return  # C-3（段）は未実装。scope フィールドだけ多重化に備えて先に定義しておく
        if not request_id or request_id == self._proposal_last_request_id:
            return  # ポーリングでの重複取得対策
        self._proposal_last_request_id = request_id

        watch = (self._payload.get("meta") or {}).get("watch") if self._payload else None
        if not watch:
            return

        if scope == "id":
            # F5: ID採番の計算は同期スキャンで近似的に即時完了するため、
            # 計算中のキャンセル対象となる長時間処理が無い（"cancel"は何もしない）。
            if action == "compute":
                self._proposal_inflight_request_id = request_id
                threading.Thread(
                    target=self._compute_id_proposal_worker,
                    args=(request_id, watch),
                    daemon=True,
                ).start()
            elif action == "update":
                self._proposal_inflight_request_id = request_id
                threading.Thread(
                    target=self._update_id_proposal_worker,
                    args=(request_id, watch),
                    daemon=True,
                ).start()
            return

        if action == "compute":
            self._proposal_stop_event.clear()
            self._proposal_stale_retries = 0
            self._proposal_inflight_request_id = request_id
            threading.Thread(
                target=self._compute_proposal_worker,
                args=(request_id, watch, mode),
                daemon=True,
            ).start()
        elif action == "update":
            self._proposal_inflight_request_id = request_id
            threading.Thread(
                target=self._update_proposal_worker,
                args=(request_id, watch, mode),
                daemon=True,
            ).start()
        elif action == "cancel":
            self._proposal_stop_event.set()

    def _run_proposal_compute(
        self, watch: dict[str, Any], mode: str
    ) -> tuple[ProposalResult, ProposalSnapshot]:
        """Excel を再読取し「提案の計算」を実行する（呼び出し側スレッドで COM 初期化する）。"""
        pythoncom.CoInitialize()
        try:
            data, _ = read_watched_range(watch)
            nodes, _, _ = parse_table_rows(data)
            raw_levels = {
                n["id"]: parse_level_optional(
                    data[n["ridx"]][LEVEL_COL] if len(data[n["ridx"]]) > LEVEL_COL else None
                )
                for n in nodes
            }
            result = compute_level_proposals(
                nodes, raw_levels, mode=mode, stop_event=self._proposal_stop_event
            )
            snapshot = capture_snapshot(data)
            return result, snapshot
        finally:
            pythoncom.CoUninitialize()

    def _compute_proposal_worker(self, request_id: str, watch: dict[str, Any], mode: str) -> None:
        try:
            result, snapshot = self._run_proposal_compute(watch, mode)
        except Exception as exc:
            logger.exception("proposal_compute_failed")
            message = str(exc)
            self._schedule_after(0, lambda: self._push_proposal_error(request_id, message))
            return

        self._proposal_baseline = snapshot
        self._proposal_result = result
        self._schedule_after(0, lambda: self._push_proposal_result(request_id, result))

    def _update_proposal_worker(self, request_id: str, watch: dict[str, Any], mode: str) -> None:
        if self._proposal_result is None or self._proposal_baseline is None:
            self._schedule_after(
                0,
                lambda: self._push_update_result(
                    request_id, ok=False, error="先に「提案の計算」を実行してください。"
                ),
            )
            return

        reminder = autosave_reminder_for_watch(watch)
        if reminder and self._on_autosave_reminder:
            self._schedule_after(0, lambda: self._on_autosave_reminder(reminder))

        try:
            write_result = write_level_updates(
                watch,
                self._proposal_result.proposals,
                mode=mode,
                baseline=self._proposal_baseline,
            )
        except Exception as exc:
            logger.exception("proposal_update_failed")
            message = str(exc)
            self._schedule_after(
                0, lambda: self._push_update_result(request_id, ok=False, error=message)
            )
            return

        decision, self._proposal_stale_retries = decide_proposal_next_action(
            write_result, self._proposal_stale_retries, max_retries=_PROPOSAL_STALE_RETRY_MAX
        )

        if decision == "give_up":
            self._schedule_after(
                0,
                lambda: self._push_update_result(
                    request_id,
                    ok=False,
                    error=(
                        "表が編集され続けているため、編集を一時停止してから"
                        "更新をやり直してください。"
                    ),
                ),
            )
            return

        if decision == "recompute":
            # 区別①（トポロジー変更）検知 → 自動再計算し、一覧を再表示してから改めて確認を求める
            try:
                result, snapshot = self._run_proposal_compute(watch, mode)
            except Exception as exc:
                logger.exception("proposal_recompute_after_stale_failed")
                message = str(exc)
                self._schedule_after(
                    0, lambda: self._push_update_result(request_id, ok=False, error=message)
                )
                return
            self._proposal_baseline = snapshot
            self._proposal_result = result
            self._schedule_after(0, lambda: self._push_proposal_result(request_id, result))
            self._schedule_after(
                0,
                lambda: self._push_update_result(request_id, ok=False, stale_topology=True),
            )
            return

        # decision == "push_result"
        self._schedule_after(
            0,
            lambda: self._push_update_result(
                request_id,
                ok=write_result.ok,
                updated_count=write_result.updated_count,
                excluded_count=write_result.excluded_count,
                error=write_result.error,
            ),
        )

    def _run_id_proposal_compute(
        self, watch: dict[str, Any]
    ) -> tuple[IdProposalResult, ProposalSnapshot, int]:
        """F5: Excel を再読取し「ID採番の提案」を計算する（呼び出し側スレッドでCOM初期化する）。"""
        pythoncom.CoInitialize()
        try:
            data, _ = read_watched_range(watch)
            result = compute_id_proposals(data)
            snapshot = capture_snapshot(data)
            return result, snapshot, len(data)
        finally:
            pythoncom.CoUninitialize()

    def _compute_id_proposal_worker(self, request_id: str, watch: dict[str, Any]) -> None:
        try:
            result, snapshot, row_count = self._run_id_proposal_compute(watch)
        except Exception as exc:
            logger.exception("id_proposal_compute_failed")
            message = str(exc)
            self._schedule_after(0, lambda: self._push_proposal_error(request_id, message))
            return

        self._id_proposal_baseline = snapshot
        self._id_proposal_result = result
        self._id_proposal_row_count = row_count
        self._schedule_after(0, lambda: self._push_id_proposal_result(request_id, result))

    def _update_id_proposal_worker(self, request_id: str, watch: dict[str, Any]) -> None:
        if (
            self._id_proposal_result is None
            or self._id_proposal_baseline is None
            or self._id_proposal_row_count is None
        ):
            self._schedule_after(
                0,
                lambda: self._push_update_result(
                    request_id, ok=False, error="先に「ID採番の計算」を実行してください。"
                ),
            )
            return

        reminder = autosave_reminder_for_watch(watch)
        if reminder and self._on_autosave_reminder:
            self._schedule_after(0, lambda: self._on_autosave_reminder(reminder))

        try:
            write_result = write_id_updates(
                watch,
                self._id_proposal_result.proposals,
                baseline_row_count=self._id_proposal_row_count,
                baseline=self._id_proposal_baseline,
            )
        except Exception as exc:
            logger.exception("id_proposal_update_failed")
            message = str(exc)
            self._schedule_after(
                0, lambda: self._push_update_result(request_id, ok=False, error=message)
            )
            return

        decision, self._id_stale_retries = decide_proposal_next_action(
            write_result, self._id_stale_retries, max_retries=_PROPOSAL_STALE_RETRY_MAX
        )

        if decision == "give_up":
            self._schedule_after(
                0,
                lambda: self._push_update_result(
                    request_id,
                    ok=False,
                    error=(
                        "表が編集され続けているため、編集を一時停止してから"
                        "更新をやり直してください。"
                    ),
                ),
            )
            return

        if decision == "recompute":
            try:
                result, snapshot, row_count = self._run_id_proposal_compute(watch)
            except Exception as exc:
                logger.exception("id_proposal_recompute_after_stale_failed")
                message = str(exc)
                self._schedule_after(
                    0, lambda: self._push_update_result(request_id, ok=False, error=message)
                )
                return
            self._id_proposal_baseline = snapshot
            self._id_proposal_result = result
            self._id_proposal_row_count = row_count
            self._schedule_after(0, lambda: self._push_id_proposal_result(request_id, result))
            self._schedule_after(
                0,
                lambda: self._push_update_result(request_id, ok=False, stale_topology=True),
            )
            return

        # decision == "push_result"
        self._schedule_after(
            0,
            lambda: self._push_update_result(
                request_id,
                ok=write_result.ok,
                updated_count=write_result.updated_count,
                excluded_count=write_result.excluded_count,
                error=write_result.error,
            ),
        )

    def _push_id_proposal_result(self, request_id: str, result: IdProposalResult) -> None:
        """F5: 計算結果をReact側（ProposalPanel scope="id"）へ`setProposalResult`で渡す。

        `ProposalResultPayload`と同じワイヤ形式に合わせ、ダングリング参照の警告は
        既存の「要確認」欄（`needsReview`）に乗せて表示する（新規の通知チャネルを
        増やさない）。
        """
        self._clear_proposal_inflight(request_id)
        payload: dict[str, Any] = {
            "requestId": request_id,
            "proposals": [
                {
                    "nodeId": f"行{p.row_index + 1}",
                    "current": None,
                    "proposed": p.proposed_id,
                    "reason": p.reason,
                }
                for p in result.proposals
            ],
            "needsReview": [{"nodeId": "-", "reason": w} for w in result.dangling_warnings],
            "skippedMultiDest": [],
        }
        self._call_js("setProposalResult", payload)

    def _push_proposal_result(self, request_id: str, result: ProposalResult) -> None:
        self._clear_proposal_inflight(request_id)
        if result.cancelled:
            payload: dict[str, Any] = {"requestId": request_id, "cancelled": True}
        else:
            payload = {
                "requestId": request_id,
                "proposals": [
                    {
                        "nodeId": p.node_id,
                        "current": p.current,
                        "proposed": p.proposed,
                        "reason": p.reason,
                    }
                    for p in result.proposals
                ],
                "needsReview": [
                    {"nodeId": r.node_id, "reason": r.reason} for r in result.needs_review
                ],
                "skippedMultiDest": list(result.skipped_multi_dest),
            }
        self._call_js("setProposalResult", payload)

    def _push_proposal_error(self, request_id: str, message: str) -> None:
        self._clear_proposal_inflight(request_id)
        self._call_js("setProposalResult", {"requestId": request_id, "error": message})

    def _clear_proposal_inflight(self, request_id: str) -> None:
        """F8: 応答を送った request_id の無応答タイムアウト対象フラグを解除する。"""
        if request_id == self._proposal_inflight_request_id:
            self._proposal_inflight_request_id = None

    def _push_update_result(
        self,
        request_id: str,
        *,
        ok: bool,
        updated_count: int = 0,
        excluded_count: int = 0,
        stale_topology: bool = False,
        error: str | None = None,
    ) -> None:
        self._clear_proposal_inflight(request_id)
        self._call_js(
            "setProposalUpdateResult",
            {
                "requestId": request_id,
                "ok": ok,
                "updatedCount": updated_count,
                "excludedCount": excluded_count,
                "staleTopology": stale_topology,
                "error": error,
            },
        )

    def _call_js(self, fn_name: str, payload: Any) -> None:
        if not self._core_ready:
            return
        js = build_set_global_js(fn_name, payload)
        try:
            core = getattr(self._frame, "core", None)
            if core is None:
                core = getattr(self._frame.web, "CoreWebView2", None)
            if core is not None:
                core.ExecuteScriptAsync(js)
            else:
                self._frame.evaluate_js(js)
        except Exception as exc:  # noqa: BLE001
            logger.warning("embedded_proposal_push_failed | fn=%s | %s", fn_name, exc)

    def _sync_webview_size(self) -> None:
        try:
            from tkwebview2.tkwebview2 import user32

            w = max(1, self._frame.winfo_width())
            h = max(1, self._frame.winfo_height())
            user32.MoveWindow(int(self._frame.chwnd), 0, 0, w, h, True)
        except Exception as exc:  # noqa: BLE001
            logger.debug("embedded_webview_resize_skip | %s", exc)

    def _on_core_ready_clr(self, sender: Any, _args: Any) -> None:
        """WinForms/CLR スレッド。Tk・winfo・after は呼ばない。"""
        self._core_ready = True
        self._pending_main_work = True

    def _on_page_loaded_bg(self) -> None:
        """pywebview loaded は別スレッドで発火しうる。Tk は触らない。"""
        self._page_loaded = True
        self._pending_main_work = True

    def _main_pump(self) -> None:
        """Tk メインスレッド専用: リサイズ・注入をここでのみ実行。"""
        if not self._pump_alive:
            return
        # core が CLR 側で後から付く場合のポーリング
        if not self._core_ready:
            core = getattr(self._frame, "core", None)
            if core is None:
                core = getattr(getattr(self._frame, "web", None), "CoreWebView2", None)
            if core is not None:
                self._core_ready = True
                self._pending_main_work = True
        if self._pending_main_work:
            self._pending_main_work = False
            self._sync_webview_size()
            if self._payload and self._core_ready:
                self._inject(self._payload)
        if self._pending_validation_result is not None:
            ok, error_count = self._pending_validation_result
            self._pending_validation_result = None
            if ok != self._validation_ok:
                self._validation_ok = ok
                logger.info(
                    "embedded_validation_reported | ok=%s | error_count=%s",
                    ok,
                    error_count,
                )
                if self._on_payload_change:
                    self._on_payload_change()
        if self._pending_proposal_action is not None:
            action_req = self._pending_proposal_action
            self._pending_proposal_action = None
            self._handle_proposal_action(action_req)
        self._schedule_after(100, self._main_pump)

    def _inject(self, payload: dict[str, Any]) -> None:
        """Core 準備後に JS 注入（Tk メインスレッドから呼ぶ）。"""
        if not self._core_ready:
            self._pending_main_work = True
            return
        js = build_payload_inject_js(payload)
        try:
            core = getattr(self._frame, "core", None)
            if core is None:
                core = getattr(self._frame.web, "CoreWebView2", None)
            if core is not None:
                core.ExecuteScriptAsync(js)
                logger.info(
                    "embedded_inject_ok | nodes=%s",
                    (payload.get("meta") or {}).get("nodeCount"),
                )
                return
            self._frame.evaluate_js(js)
            logger.info(
                "embedded_inject_ok_via_evaluate | nodes=%s",
                (payload.get("meta") or {}).get("nodeCount"),
            )
        except Exception as exc:  # noqa: BLE001
            logger.warning("embedded_inject_failed | %s", exc)
            if self._inject_retries < _INJECT_RETRY_MAX:
                self._inject_retries += 1
                self._schedule_after(_INJECT_RETRY_MS, lambda: self._inject(payload))

    def _schedule_live(self) -> None:
        if not self._live:
            return
        self._schedule_after(int(LIVE_POLL_INTERVAL_SEC * 1000), self._live_tick)

    def _live_tick(self) -> None:
        if not self._live or not self._payload:
            return
        if self._core_ready:
            self._poll_validation()
            self._poll_proposal_action()
            self._check_js_response_timeout()
        watch = (self._payload.get("meta") or {}).get("watch")
        if watch:
            previous_row_count = self._f4_row_count
            self._f4_row_count = apply_new_row_validation(
                watch,
                previous_row_count,
                type_validation_list=SHAPE_TYPE_VALIDATION,
                color_validation_list=COLOR_VALIDATION,
            )
            if (
                previous_row_count is not None
                and self._f4_row_count is not None
                and self._f4_row_count > previous_row_count
                and self._on_row_validation_applied
            ):
                self._on_row_validation_applied(self._f4_row_count - previous_row_count)
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
                self._inject_retries = 0
                self._inject(fresh)
                if self._on_payload_change:
                    self._on_payload_change()
        self._schedule_live()
