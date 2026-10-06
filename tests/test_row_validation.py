"""apply_new_row_validation のユニットテスト（構想設計§5、COM は MagicMock）。"""
import unittest
from unittest.mock import MagicMock, patch

from app.core.row_validation import apply_new_row_validation

_WATCH_FULL = {
    "workbookName": "Book1.xlsx",
    "sheetName": "Sheet1",
    "anchorAddress": "$A$1",
    "isFullMode": True,
}

_WATCH_RANGE = {
    "workbookName": "Book1.xlsx",
    "sheetName": "Sheet1",
    "anchorAddress": "$A$1",
    "rangeAddress": "$A$1:$J$50",
    "isFullMode": False,
}


def _make_app(row_count: int):
    app = MagicMock()
    workbook = MagicMock()
    workbook.Name = _WATCH_FULL["workbookName"]
    app.Workbooks = [workbook]
    sheet = MagicMock()
    workbook.Sheets.return_value = sheet
    r_tgt = MagicMock()
    r_tgt.Rows.Count = row_count
    sheet.Range.return_value.CurrentRegion = r_tgt
    return app, r_tgt


class ApplyNewRowValidationTests(unittest.TestCase):
    def test_range_mode_is_out_of_scope_and_returns_previous_unchanged(self) -> None:
        with patch("app.core.row_validation.get_excel_app") as mock_get_app:
            result = apply_new_row_validation(
                _WATCH_RANGE, 10, type_validation_list="a", color_validation_list="b"
            )
        self.assertEqual(result, 10)
        mock_get_app.assert_not_called()

    def test_first_call_only_establishes_baseline_without_applying(self) -> None:
        app, r_tgt = _make_app(row_count=105)
        with patch("app.core.row_validation.get_excel_app", return_value=app):
            result = apply_new_row_validation(
                _WATCH_FULL, None, type_validation_list="a", color_validation_list="b"
            )
        self.assertEqual(result, 105)
        r_tgt.Rows.assert_not_called()

    def test_row_count_increase_applies_validation_to_new_rows_only(self) -> None:
        app, r_tgt = _make_app(row_count=103)
        with patch("app.core.row_validation.get_excel_app", return_value=app):
            result = apply_new_row_validation(
                _WATCH_FULL, 100, type_validation_list="TYPE_LIST", color_validation_list="COLOR_LIST"
            )
        self.assertEqual(result, 103)
        r_tgt.Rows.assert_called_once_with(101)  # previous_row_count+1 = 新規行の先頭
        new_rows = r_tgt.Rows.return_value
        new_rows.Resize.assert_called_once_with(3)  # added = 103-100
        # 種別列(TYPE_COL+1=2)・色列(COLOR_COL+1=3)の両方にそれぞれの入力規則が適用される
        range_mock = new_rows.Resize.return_value.Cells.return_value.Resize.return_value
        range_mock.Validation.Add.assert_any_call(3, 1, 1, "TYPE_LIST")
        range_mock.Validation.Add.assert_any_call(3, 1, 1, "COLOR_LIST")
        self.assertEqual(range_mock.Validation.Add.call_count, 2)

    def test_row_count_unchanged_does_not_touch_validation(self) -> None:
        app, r_tgt = _make_app(row_count=100)
        with patch("app.core.row_validation.get_excel_app", return_value=app):
            result = apply_new_row_validation(
                _WATCH_FULL, 100, type_validation_list="a", color_validation_list="b"
            )
        self.assertEqual(result, 100)
        r_tgt.Rows.assert_not_called()

    def test_row_count_decrease_does_not_touch_validation(self) -> None:
        app, r_tgt = _make_app(row_count=90)
        with patch("app.core.row_validation.get_excel_app", return_value=app):
            result = apply_new_row_validation(
                _WATCH_FULL, 100, type_validation_list="a", color_validation_list="b"
            )
        self.assertEqual(result, 90)
        r_tgt.Rows.assert_not_called()

    def test_excel_not_running_returns_previous_unchanged(self) -> None:
        with patch("app.core.row_validation.get_excel_app", return_value=None):
            result = apply_new_row_validation(
                _WATCH_FULL, 100, type_validation_list="a", color_validation_list="b"
            )
        self.assertEqual(result, 100)


if __name__ == "__main__":
    unittest.main()
