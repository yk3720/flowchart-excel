"""write_id_updates のユニットテスト（構想設計§6-2、COM は MagicMock）。"""
import unittest
from unittest.mock import MagicMock, PropertyMock, patch

import pywintypes

from app.core.id_proposal import IdProposal
from app.core.id_writer import write_id_updates
from app.core.proposal_fingerprint import capture_snapshot

_WATCH = {
    "workbookName": "Book1.xlsx",
    "sheetName": "Sheet1",
    "anchorAddress": "$A$1",
    "rangeAddress": "$A$1:$J$3",
    "isFullMode": False,
}

# table-10col-v2: ID, 種別, 色, 下, 右, 段, 列, Text1, Text2, Text3
_ROW_A = ("10", "端子", "", "20", "", 0, 0, "start", "", "")
_ROW_B = ("20", "処理", "", "", "", 1, 5, "step", "", "")
_NEW_ROW_BLANK_ID = ("", "判断", "", "", "", None, None, "", "", "")


def _make_workbook_sheet(fresh_rows):
    app = MagicMock()
    workbook = MagicMock()
    workbook.Name = _WATCH["workbookName"]
    app.Workbooks = [workbook]
    sheet = MagicMock()
    workbook.Sheets.return_value = sheet

    r_tgt = MagicMock()
    sheet.Range.return_value = r_tgt
    r_tgt.Value = tuple(fresh_rows)
    r_tgt.Rows.Count = len(fresh_rows)

    id_range = MagicMock()
    r_tgt.Cells.return_value.Resize.return_value = id_range
    return app, id_range


class WriteIdUpdatesTests(unittest.TestCase):
    def test_no_proposals_short_circuits_without_excel_access(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        with patch("app.core.id_writer.read_watched_range") as mock_read:
            result = write_id_updates(
                _WATCH, (), baseline_row_count=2, baseline=baseline
            )
        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 0)
        mock_read.assert_not_called()

    def test_writes_proposed_id_into_target_row_only(self) -> None:
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        app, id_range = _make_workbook_sheet(fresh_rows)
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 1)
        written = id_range.Value
        self.assertEqual(written[2][0], "21")
        self.assertEqual(written[0][0], "10")  # 既存行のIDは不変
        self.assertEqual(written[1][0], "20")

    def test_existing_row_topology_change_blocks_write(self) -> None:
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        changed_row_a = ("10", "端子", "", "99", "", 0, 0, "start", "", "")  # 接続先(下)変更
        current_rows = (changed_row_a, _ROW_B, _NEW_ROW_BLANK_ID)
        app, _id_range = _make_workbook_sheet(current_rows)
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(current_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertFalse(result.ok)
        self.assertTrue(result.stale_topology)

    def test_row_count_changed_since_proposal_blocks_write(self) -> None:
        """提案計算時から行数が変わっていれば、他のidとtopology両方不変でも安全側で中断する。"""
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        extra_new_row = ("", "処理", "", "", "", None, None, "", "", "")
        current_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID, extra_new_row)
        app, _id_range = _make_workbook_sheet(current_rows)
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(current_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertFalse(result.ok)
        self.assertTrue(result.stale_topology)

    def test_target_row_no_longer_blank_id_blocks_write(self) -> None:
        """提案後に対象行へ手動でIDが入力された場合、書き込みを中断する。"""
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        hand_filled_row = ("99", "判断", "", "", "", None, None, "", "", "")
        current_rows = (_ROW_A, _ROW_B, hand_filled_row)
        app, _id_range = _make_workbook_sheet(current_rows)
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(current_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertFalse(result.ok)
        self.assertTrue(result.stale_topology)

    def test_com_error_during_write_is_reported_not_swallowed(self) -> None:
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        app, id_range = _make_workbook_sheet(fresh_rows)
        type(id_range).Value = PropertyMock(
            side_effect=pywintypes.com_error(-1, "write failed", None, None)
        )
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertFalse(result.ok)
        self.assertIsNotNone(result.error)

    def test_screen_updating_restored_after_com_error(self) -> None:
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        app, id_range = _make_workbook_sheet(fresh_rows)
        type(id_range).Value = PropertyMock(
            side_effect=pywintypes.com_error(-1, "write failed", None, None)
        )
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            write_id_updates(_WATCH, proposals, baseline_row_count=3, baseline=baseline)

        self.assertTrue(app.ScreenUpdating)
        self.assertTrue(app.DisplayAlerts)


if __name__ == "__main__":
    unittest.main()
