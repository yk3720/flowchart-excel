"""write_id_updates のユニットテスト（構想設計§6-2、COM は MagicMock）。"""
import unittest
from unittest.mock import MagicMock, patch

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


class _FakeCell:
    def __init__(self, rows: list[list[object]], row_1based: int, col_1based: int, persist: bool):
        self._rows = rows
        self._r = row_1based - 1
        self._c = col_1based - 1
        self._persist = persist

    @property
    def Value(self) -> object:
        return self._rows[self._r][self._c]

    @Value.setter
    def Value(self, value: object) -> None:
        if self._persist:
            self._rows[self._r][self._c] = value


class _FakeRange:
    def __init__(self, rows: list[list[object]], persist_writes: bool):
        self._rows = rows
        self._persist = persist_writes
        self.Rows = MagicMock()
        self.Rows.Count = len(rows)

    @property
    def Value(self) -> tuple:
        return tuple(tuple(row) for row in self._rows)

    def Cells(self, row_1based: int, col_1based: int) -> _FakeCell:
        return _FakeCell(self._rows, row_1based, col_1based, self._persist)


def _make_workbook_sheet(fresh_rows, *, persist_writes: bool = True):
    app = MagicMock()
    workbook = MagicMock()
    workbook.Name = _WATCH["workbookName"]
    app.Workbooks = [workbook]
    sheet = MagicMock()
    workbook.Sheets.return_value = sheet

    rows = [list(row) for row in fresh_rows]
    r_tgt = _FakeRange(rows, persist_writes)
    sheet.Range.return_value = r_tgt
    return app, r_tgt, rows


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
        app, _r_tgt, rows = _make_workbook_sheet(fresh_rows)
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertTrue(result.ok)
        self.assertEqual(result.updated_count, 1)
        self.assertEqual(rows[2][0], "21")
        self.assertEqual(rows[0][0], "10")  # 既存行のIDは不変
        self.assertEqual(rows[1][0], "20")

    def test_existing_row_topology_change_blocks_write(self) -> None:
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        changed_row_a = ("10", "端子", "", "99", "", 0, 0, "start", "", "")  # 接続先(下)変更
        current_rows = (changed_row_a, _ROW_B, _NEW_ROW_BLANK_ID)
        app, _r_tgt, _rows = _make_workbook_sheet(current_rows)
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
        app, _r_tgt, _rows = _make_workbook_sheet(current_rows)
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
        app, _r_tgt, _rows = _make_workbook_sheet(current_rows)
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
        app, r_tgt, _rows = _make_workbook_sheet(fresh_rows)

        def boom(_row, _col):
            raise pywintypes.com_error(-1, "write failed", None, None)

        r_tgt.Cells = boom  # type: ignore[method-assign]
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertFalse(result.ok)
        self.assertIsNotNone(result.error)

    def test_verify_failure_when_excel_does_not_reflect_write(self) -> None:
        """セル代入が例外なしでも反映されない場合は成功扱いにしない。"""
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        app, _r_tgt, rows = _make_workbook_sheet(fresh_rows, persist_writes=False)
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            result = write_id_updates(
                _WATCH, proposals, baseline_row_count=3, baseline=baseline
            )

        self.assertFalse(result.ok)
        self.assertIn("反映されませんでした", result.error or "")
        self.assertEqual(rows[2][0], "")  # 実データは空のまま

    def test_screen_updating_restored_after_com_error(self) -> None:
        fresh_rows = (_ROW_A, _ROW_B, _NEW_ROW_BLANK_ID)
        baseline = capture_snapshot(fresh_rows)
        app, r_tgt, _rows = _make_workbook_sheet(fresh_rows)

        def boom(_row, _col):
            raise pywintypes.com_error(-1, "write failed", None, None)

        r_tgt.Cells = boom  # type: ignore[method-assign]
        proposals = (IdProposal(row_index=2, proposed_id="21", reason="test"),)

        with patch("app.core.id_writer.read_watched_range", return_value=(fresh_rows, "t")), \
             patch("app.core.id_writer.get_excel_app", return_value=app):
            write_id_updates(_WATCH, proposals, baseline_row_count=3, baseline=baseline)

        self.assertTrue(app.ScreenUpdating)
        self.assertTrue(app.DisplayAlerts)


if __name__ == "__main__":
    unittest.main()
