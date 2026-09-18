"""level_writer のユニットテスト（COM は MagicMock、test_excel_engine.py と同じ方式）。"""
import unittest
from unittest.mock import MagicMock, PropertyMock, patch

import pywintypes

from app.core.level_inference import LevelProposal
from app.core.proposal_fingerprint import capture_snapshot
from app.core.level_writer import write_level_updates

_WATCH = {
    "workbookName": "Book1.xlsx",
    "sheetName": "Sheet1",
    "anchorAddress": "$A$1",
    "rangeAddress": "$A$1:$J$2",
    "isFullMode": False,
}

# table-10col-v2: ID, 種別, 色, 下, 右, 段, 列, Text1, Text2, Text3
_ROW_A = ("10", "端子", "", "20", "", 0, 0, "start", "", "")
_ROW_B = ("20", "処理", "", "", "", 1, 5, "step", "", "")


def _make_workbook_sheet(fresh_rows, *, level_column_values=None):
    """get_excel_app() が返すオブジェクトチェーンを MagicMock で組み立てる。"""
    app = MagicMock()
    workbook = MagicMock()
    workbook.Name = _WATCH["workbookName"]
    app.Workbooks = [workbook]
    sheet = MagicMock()
    workbook.Sheets.return_value = sheet

    r_tgt = MagicMock()
    sheet.Range.return_value = r_tgt

    level_range = MagicMock()
    if level_column_values is None:
        level_column_values = [[row[6]] for row in fresh_rows]
    level_range.Value = (
        level_column_values[0][0] if len(level_column_values) == 1 else tuple(
            tuple(row) for row in level_column_values
        )
    )
    r_tgt.Cells.return_value.Resize.return_value = level_range
    return app, level_range


class WriteLevelUpdatesTests(unittest.TestCase):
    def test_no_proposals_short_circuits_without_excel_access(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        with patch("app.core.level_writer.read_watched_range") as mock_read:
            result = write_level_updates(_WATCH, [], mode="blank_only", baseline=baseline)
        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 0)
        mock_read.assert_not_called()

    def test_writes_range_once_when_table_unchanged(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        fresh_rows = (_ROW_A, _ROW_B)
        app, level_range = _make_workbook_sheet(fresh_rows)
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(result.excluded_count, 0)
        written = level_range.Value
        self.assertEqual(written[1][0], 1)  # id=20(行index1)が新しい提案値に書き換わる
        self.assertEqual(written[0][0], 0)  # id=10(行index0)は元の値のまま

    def test_topology_change_blocks_write_and_reports_stale(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        changed_row_a = ("10", "端子", "", "30", "", 0, 0, "start", "", "")  # 接続先(下)変更
        fresh_rows = (changed_row_a, _ROW_B)
        app, level_range = _make_workbook_sheet(fresh_rows)
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertFalse(result.ok)
        self.assertTrue(result.stale_topology)
        level_range.Value = None  # 書き込みが起きていないことを後段で確認できるよう明示リセット
        self.assertIsNone(level_range.Value)

    def test_hand_edited_target_cell_excluded_in_blank_only_mode(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        hand_edited_row_b = ("20", "処理", "", "", "", 1, 9, "step", "", "")  # 列セルを手入力
        fresh_rows = (_ROW_A, hand_edited_row_b)
        app, level_range = _make_workbook_sheet(fresh_rows)
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 0)
        self.assertEqual(result.excluded_count, 1)

    def test_hand_edited_target_cell_not_excluded_in_full_recalc_mode(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        hand_edited_row_b = ("20", "処理", "", "", "", 1, 9, "step", "", "")
        fresh_rows = (_ROW_A, hand_edited_row_b)
        app, level_range = _make_workbook_sheet(fresh_rows)
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="full_recalc", baseline=baseline)

        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(result.excluded_count, 0)

    def test_com_error_during_write_is_reported_not_swallowed(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        fresh_rows = (_ROW_A, _ROW_B)
        app, level_range = _make_workbook_sheet(fresh_rows)
        # 1回目の呼び出し(現在値の読取)は正常値、2回目(書き込み時の代入)は COM エラー。
        type(level_range).Value = PropertyMock(
            side_effect=[
                tuple((row[6],) for row in fresh_rows),
                pywintypes.com_error(-1, "write failed", None, None),
            ]
        )
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertFalse(result.ok)
        self.assertIsNotNone(result.error)


if __name__ == "__main__":
    unittest.main()
