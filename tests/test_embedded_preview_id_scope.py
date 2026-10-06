"""F5（scope="id"）の_handle_proposal_action〜pushまでの結合テスト。

`EmbeddedStudioPreview.__init__`は実WebView2初期化を要求するため、
`object.__new__`でコンストラクタを経由せずインスタンスを作り、
このテストに必要な属性だけを手動で設定する（test_embedded_preview_timeout.py
と同じ方式）。
"""
import time
import unittest
from unittest.mock import MagicMock, patch

from app.ui.embedded_preview import EmbeddedStudioPreview

# table-10col-v2: ID, 種別, 色, 下, 右, 段, 列, Text1, Text2, Text3
_ROW_A = ("10", "端子", "", "20", "", 0, 0, "start", "", "")
_NEW_ROW = ("", "処理", "", "", "", None, None, "", "", "")
_WATCH = {
    "workbookName": "Book1.xlsx",
    "sheetName": "Sheet1",
    "anchorAddress": "$A$1",
    "isFullMode": True,
}


def _make_bare_preview() -> EmbeddedStudioPreview:
    preview = object.__new__(EmbeddedStudioPreview)
    preview._core_ready = False
    preview._payload = {"meta": {"watch": _WATCH}}
    preview._proposal_last_request_id = None
    preview._proposal_stop_event = MagicMock()
    preview._proposal_inflight_request_id = None
    preview._id_proposal_baseline = None
    preview._id_proposal_result = None
    preview._id_proposal_row_count = None
    preview._id_stale_retries = 0
    preview._on_autosave_reminder = None
    # 実際の `self.after` ではなく即時実行（テスト用）。
    preview._schedule_after = lambda _delay, fn: fn()
    return preview


def _wait_for(predicate, timeout=2.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return
        time.sleep(0.01)
    raise AssertionError("timed out waiting for background worker")


class IdScopeIntegrationTests(unittest.TestCase):
    def test_compute_then_update_round_trip(self) -> None:
        preview = _make_bare_preview()
        pushed_results = []
        pushed_updates = []
        preview._call_js = lambda fn_name, payload: (
            pushed_results.append(payload)
            if fn_name == "setProposalResult"
            else pushed_updates.append(payload)
        )

        fresh_rows = (_ROW_A, _NEW_ROW)
        with patch(
            "app.ui.embedded_preview.read_watched_range", return_value=(fresh_rows, "t")
        ):
            preview._handle_proposal_action(
                {"requestId": "r1", "scope": "id", "action": "compute", "mode": "blank_only"}
            )
            _wait_for(lambda: len(pushed_results) == 1)

        self.assertEqual(pushed_results[0]["requestId"], "r1")
        self.assertEqual(len(pushed_results[0]["proposals"]), 1)
        self.assertEqual(pushed_results[0]["proposals"][0]["proposed"], "11")
        self.assertIsNotNone(preview._id_proposal_result)
        self.assertEqual(preview._id_proposal_row_count, 2)

        app = MagicMock()
        workbook = MagicMock()
        workbook.Name = _WATCH["workbookName"]
        app.Workbooks = [workbook]
        sheet = MagicMock()
        workbook.Sheets.return_value = sheet
        r_tgt = MagicMock()
        sheet.Range.return_value.CurrentRegion = r_tgt  # isFullMode=True 用
        r_tgt.Value = fresh_rows
        r_tgt.Rows.Count = 2

        with patch(
            "app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")
        ), patch("app.core.id_writer.get_excel_app", return_value=app), patch(
            "app.core.autosave_check.get_excel_app", return_value=None
        ):
            preview._handle_proposal_action(
                {"requestId": "r2", "scope": "id", "action": "update", "mode": "blank_only"}
            )
            _wait_for(lambda: len(pushed_updates) == 1)

        self.assertEqual(pushed_updates[0]["requestId"], "r2")
        self.assertTrue(pushed_updates[0]["ok"])
        self.assertEqual(pushed_updates[0]["updatedCount"], 1)
        written = r_tgt.Cells.return_value.Resize.return_value.Value
        self.assertEqual(written[1][0], "11")

    def test_level_scope_action_is_unaffected_by_id_scope_wiring(self) -> None:
        """scopeディスパッチ追加が既存のlevelスコープ経路を壊していないことを確認。"""
        preview = _make_bare_preview()
        preview._proposal_stale_retries = 0
        pushed_errors = []
        preview._push_proposal_error = lambda request_id, message: pushed_errors.append(
            (request_id, message)
        )

        with patch(
            "app.ui.embedded_preview.read_watched_range", side_effect=RuntimeError("boom")
        ):
            preview._handle_proposal_action(
                {"requestId": "r3", "scope": "level", "action": "compute", "mode": "blank_only"}
            )
            _wait_for(lambda: len(pushed_errors) == 1)

        self.assertEqual(pushed_errors[0][0], "r3")


if __name__ == "__main__":
    unittest.main()
