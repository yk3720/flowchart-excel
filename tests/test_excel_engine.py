"""ExcelFlowchartEngine._draw_core のロールバック挙動のユニットテスト。"""
import threading
import unittest
from unittest.mock import MagicMock, patch

from app.core.excel_engine import ExcelFlowchartEngine

DATA = (
    ("10", "処理", "", "20", "", 0, 0, "A", "", ""),
    ("20", "処理", "", "", "", 1, 0, "B", "", ""),
)
CONFIG = {"width": 160.0, "height": 60.0, "gap_v": 30.0, "gap_h": 100.0}
THEME = {"shape_line": 0, "connector": 0}


def _make_sheet() -> MagicMock:
    sheet = MagicMock()
    sheet.Shapes.AddShape.return_value.Height = 60.0
    # sheet.Shapes(name) は名前ごとに別々のモックを返す（実際のExcelの挙動に合わせる）
    shape_by_name: dict = {}
    sheet.Shapes.side_effect = lambda name: shape_by_name.setdefault(name, MagicMock())
    return sheet


class DrawCoreRollbackTests(unittest.TestCase):
    def _engine_and_context(self):
        stop_event = threading.Event()
        engine = ExcelFlowchartEngine(stop_event)
        sheet = _make_sheet()
        start_cell = MagicMock(Left=0.0, Top=0.0)
        return engine, sheet, start_cell, stop_event

    def test_rollback_on_cancel_before_connect(self) -> None:
        engine, sheet, start_cell, stop_event = self._engine_and_context()

        def _place_shapes(sheet_, row_map, row_heights, *a, **k):
            stop_event.set()  # place_shapes の途中でユーザーが中止した想定
            return ({}, ["s1", "s2"], [], (0, 0, 10, 10))

        with patch("app.core.excel_engine.get_excel_app", return_value=MagicMock()), \
             patch("app.core.excel_engine.place_shapes", side_effect=_place_shapes), \
             patch("app.core.excel_engine.connect_nodes") as mock_connect:
            result = engine._draw_core(
                data=DATA, sheet=sheet, start_cell=start_cell, title_txt="t",
                is_full_mode=False, config=CONFIG, theme=THEME,
            )

        self.assertEqual(result, "")
        mock_connect.assert_not_called()
        sheet.Shapes("s1").Delete.assert_called_once()
        sheet.Shapes("s2").Delete.assert_called_once()

    def test_rollback_on_exception_during_finalize(self) -> None:
        engine, sheet, start_cell, _stop_event = self._engine_and_context()

        with patch("app.core.excel_engine.get_excel_app", return_value=MagicMock()), \
             patch(
                 "app.core.excel_engine.place_shapes",
                 return_value=({}, ["s1"], [], (0, 0, 10, 10)),
             ), \
             patch("app.core.excel_engine.connect_nodes", return_value=["c1"]), \
             patch(
                 "app.core.excel_engine.finalize_composites",
                 side_effect=RuntimeError("boom"),
             ):
            with self.assertRaises(RuntimeError):
                engine._draw_core(
                    data=DATA, sheet=sheet, start_cell=start_cell, title_txt="t",
                    is_full_mode=False, config=CONFIG, theme=THEME,
                )

        sheet.Shapes("s1").Delete.assert_called_once()
        sheet.Shapes("c1").Delete.assert_called_once()

    def test_final_grouping_failure_rolls_back_and_propagates(self) -> None:
        engine, sheet, start_cell, _stop_event = self._engine_and_context()

        with patch("app.core.excel_engine.get_excel_app", return_value=MagicMock()), \
             patch(
                 "app.core.excel_engine.place_shapes",
                 return_value=({}, ["s1"], [], (0, 0, 10, 10)),
             ), \
             patch("app.core.excel_engine.connect_nodes", return_value=["c1"]), \
             patch("app.core.excel_engine.finalize_composites", return_value=[]), \
             patch(
                 "app.core.excel_engine.create_final_groups",
                 side_effect=RuntimeError("group fail"),
             ):
            with self.assertRaises(RuntimeError):
                engine._draw_core(
                    data=DATA, sheet=sheet, start_cell=start_cell, title_txt="t",
                    is_full_mode=False, config=CONFIG, theme=THEME,
                )

        sheet.Shapes("s1").Delete.assert_called_once()
        sheet.Shapes("c1").Delete.assert_called_once()


if __name__ == "__main__":
    unittest.main()
