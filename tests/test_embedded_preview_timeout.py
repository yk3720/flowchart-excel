"""EmbeddedStudioPreview の無応答タイムアウト検出（構想設計§4）のユニットテスト。

`EmbeddedStudioPreview.__init__` は実際の WebView2/CLR 初期化を要求するため、
`object.__new__` でコンストラクタを経由せずインスタンスを作り、タイムアウト判定に
必要な属性だけを手動で設定する（`_core_ready=False` なら `_call_js` は何もせず
即returnするため、実WebViewなしで安全に呼べる）。
"""
import time
import unittest

from app.ui.embedded_preview import (
    _JS_RESPONSE_TIMEOUT_SEC,
    EmbeddedStudioPreview,
)


def _make_bare_preview() -> EmbeddedStudioPreview:
    preview = object.__new__(EmbeddedStudioPreview)
    preview._core_ready = False
    preview._last_js_response_at = time.monotonic()
    preview._js_timeout_notified = False
    preview._proposal_inflight_request_id = None
    preview._on_js_timeout = None
    return preview


class JsResponseTimeoutTests(unittest.TestCase):
    def test_not_timed_out_when_response_recent(self) -> None:
        preview = _make_bare_preview()
        preview._check_js_response_timeout()
        self.assertFalse(preview._js_timeout_notified)

    def test_timed_out_notifies_once(self) -> None:
        preview = _make_bare_preview()
        preview._last_js_response_at = time.monotonic() - (_JS_RESPONSE_TIMEOUT_SEC + 1)
        calls = []
        preview._on_js_timeout = lambda: calls.append(1)

        preview._check_js_response_timeout()
        preview._check_js_response_timeout()
        preview._check_js_response_timeout()

        self.assertTrue(preview._js_timeout_notified)
        self.assertEqual(len(calls), 1)  # 応答が戻るまで繰り返し通知しない

    def test_fresh_response_clears_notified_flag(self) -> None:
        preview = _make_bare_preview()
        preview._last_js_response_at = time.monotonic() - (_JS_RESPONSE_TIMEOUT_SEC + 1)
        preview._check_js_response_timeout()
        self.assertTrue(preview._js_timeout_notified)

        preview._last_js_response_at = time.monotonic()  # 応答が戻った
        preview._check_js_response_timeout()
        self.assertFalse(preview._js_timeout_notified)

    def test_handle_timeout_pushes_result_for_inflight_request_and_clears_it(self) -> None:
        preview = _make_bare_preview()
        preview._proposal_inflight_request_id = "req-1"

        preview._handle_js_response_timeout()

        self.assertIsNone(preview._proposal_inflight_request_id)

    def test_handle_timeout_without_inflight_request_does_not_crash(self) -> None:
        preview = _make_bare_preview()
        preview._handle_js_response_timeout()  # 例外が出ないことを確認
        self.assertIsNone(preview._proposal_inflight_request_id)


if __name__ == "__main__":
    unittest.main()
