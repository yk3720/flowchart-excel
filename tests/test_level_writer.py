"""level_writer のユニットテスト（COM は MagicMock、test_excel_engine.py と同じ方式）。"""
import unittest
from unittest.mock import MagicMock, PropertyMock, patch

import pywintypes

from app.core.level_inference import LevelProposal
from app.core.level_writer import write_level_updates
from app.core.proposal_fingerprint import capture_snapshot

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


def _make_workbook_sheet(fresh_rows):
    """get_excel_app() が返すオブジェクトチェーンを MagicMock で組み立てる。

    `r_tgt.Value` が「統一読み取り」（ID列・level列を含む全列を1回のCOM呼び出し
    で取得する処理）に対応する。書き込み先 `level_range`
    （`r_tgt.Cells(...).Resize(...)`）は`.Value`への代入（書き込み）にのみ使い、
    読み取りには使わない（構想設計§1-1ステップ4: 書き込み直前の2回目の読み取りを
    行わない）。
    """
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

    level_range = MagicMock()
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

    def test_row_added_between_freshness_check_and_write_blocks_write(self) -> None:
        """区別①のbaseline比較をすり抜けた後、統一読み取り時点でID集合が変わっているケース。

        baselineとの比較(topology_changed)に使う`data`（鮮度チェック用の読み取り）と、
        ワークブック/シート解決後の統一読み取り(`r_tgt.Value`)が異なるタイミングの
        COM呼び出しであるため、両者の間で行が追加されるとこの経路でしか検知できない
        （構想設計§1-1ステップ2のID集合チェック）。
        """
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        fresh_rows_at_check = (_ROW_A, _ROW_B)  # topology_changed はこちらと比較し素通りする
        row_c = ("30", "処理", "", "", "", 1, 5, "step", "", "")
        fresh_rows_at_unified_read = (_ROW_A, _ROW_B, row_c)  # 統一読み取り時点で1行増えている
        app, _level_range = _make_workbook_sheet(fresh_rows_at_unified_read)
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch(
            "app.core.level_writer.read_watched_range",
            return_value=(fresh_rows_at_check, "t"),
        ), patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertFalse(result.ok)
        self.assertTrue(result.stale_topology)

    def test_hand_edited_target_cell_excluded_in_blank_only_mode(self) -> None:
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        hand_edited_row_b = ("20", "処理", "", "", "", 1, 9, "step", "", "")  # 列セルを手入力
        fresh_rows = (_ROW_A, hand_edited_row_b)
        app, _level_range = _make_workbook_sheet(fresh_rows)
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
        app, _level_range = _make_workbook_sheet(fresh_rows)
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
        # 統一読み取り(r_tgt.Value)は正常値、書き込み(level_range.Value への代入)でCOMエラー。
        type(level_range).Value = PropertyMock(
            side_effect=pywintypes.com_error(-1, "write failed", None, None)
        )
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            result = write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertFalse(result.ok)
        self.assertIsNotNone(result.error)

    def test_screen_updating_restored_after_com_error(self) -> None:
        """§1-3: 書き込み失敗時でも ScreenUpdating/DisplayAlerts が必ず復帰する。"""
        baseline = capture_snapshot((_ROW_A, _ROW_B))
        fresh_rows = (_ROW_A, _ROW_B)
        app, level_range = _make_workbook_sheet(fresh_rows)
        type(level_range).Value = PropertyMock(
            side_effect=pywintypes.com_error(-1, "write failed", None, None)
        )
        proposals = [LevelProposal(node_id="20", current=5, proposed=1, reason="test")]

        with patch("app.core.level_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.level_writer.get_excel_app", return_value=app):
            write_level_updates(_WATCH, proposals, mode="blank_only", baseline=baseline)

        self.assertTrue(app.ScreenUpdating)
        self.assertTrue(app.DisplayAlerts)


if __name__ == "__main__":
    unittest.main()
