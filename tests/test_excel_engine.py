"""ExcelFlowchartEngine._draw_core のロールバック挙動のユニットテスト。"""
import threading
import unittest
from unittest.mock import MagicMock, patch

from app.core.excel_engine import ExcelFlowchartEngine
from app.core.group_manager import frame_anchor_offset

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

    def test_full_mode_offsets_anchor_so_frame_corner_matches_selected_cell(self) -> None:
        """表全体モードでは、外枠の角が選択セルに一致するよう図形群の原点をずらす。"""
        engine, sheet, _start_cell, _stop_event = self._engine_and_context()
        start_cell = MagicMock(Left=100.0, Top=200.0)

        with patch("app.core.excel_engine.get_excel_app", return_value=MagicMock()), \
             patch(
                 "app.core.excel_engine.place_shapes",
                 return_value=({}, ["s1"], [], (0, 0, 10, 10)),
             ) as mock_place_shapes, \
             patch("app.core.excel_engine.connect_nodes", return_value=[]), \
             patch("app.core.excel_engine.finalize_composites", return_value=[]), \
             patch("app.core.excel_engine.add_frame_and_title", return_value=[]), \
             patch("app.core.excel_engine.create_final_groups", return_value="grp"):
            engine._draw_core(
                data=DATA, sheet=sheet, start_cell=start_cell, title_txt="t",
                is_full_mode=True, config=CONFIG, theme=THEME,
            )

        called_left = mock_place_shapes.call_args.args[3]
        called_top = mock_place_shapes.call_args.args[4]
        dx, dy = frame_anchor_offset()
        self.assertEqual(called_left, 100.0 + dx)
        self.assertEqual(called_top, 200.0 + dy)


class ReadCurrentAnchorTests(unittest.TestCase):
    def test_uses_active_selection_not_stored_watch(self) -> None:
        """描画位置は実行時点の選択セルを使う（プレビュー起動時に覚えたwatchは使わない）。"""
        engine = ExcelFlowchartEngine(threading.Event())
        app = MagicMock()
        current_cell = MagicMock()
        app.Selection.Cells.return_value = current_cell
        app.ActiveSheet = "current_sheet"

        with patch("app.core.excel_engine.get_excel_app", return_value=app):
            sheet, start_cell = engine._read_current_anchor()

        self.assertEqual(sheet, "current_sheet")
        self.assertIs(start_cell, current_cell)
        app.Selection.Cells.assert_called_once_with(1, 1)

    def test_raises_when_excel_not_running(self) -> None:
        engine = ExcelFlowchartEngine(threading.Event())
        with patch("app.core.excel_engine.get_excel_app", return_value=None):
            with self.assertRaises(RuntimeError):
                engine._read_current_anchor()


class DrawFromStudioPayloadAnchorTests(unittest.TestCase):
    def test_ignores_stale_watch_anchor_and_uses_current_selection(self) -> None:
        """payload の meta.watch.anchorAddress が古くても、実行時点の選択セルで描画すること。"""
        engine = ExcelFlowchartEngine(threading.Event())
        current_sheet = MagicMock()
        current_cell = MagicMock()

        payload = {
            "table": [["10", "処理", "", "", "", 0, 0, "A", "", ""]],
            "layout": {"width": 160.0, "heightMin": 60.0, "gapV": 30.0, "gapH": 100.0},
            "title": "t",
            "isFullMode": False,
            "meta": {
                "watch": {
                    "workbookName": "stale.xlsx",
                    "sheetName": "Sheet1",
                    "anchorAddress": "$B$5",
                }
            },
        }

        with patch(
            "app.core.excel_engine.ExcelFlowchartEngine._read_current_anchor",
            return_value=(current_sheet, current_cell),
        ) as mock_anchor, patch.object(
            engine, "_draw_core", return_value="grp"
        ) as mock_draw_core:
            engine.draw_from_studio_payload(payload, theme=THEME)

        mock_anchor.assert_called_once()
        self.assertIs(mock_draw_core.call_args.kwargs["sheet"], current_sheet)
        self.assertIs(mock_draw_core.call_args.kwargs["start_cell"], current_cell)


if __name__ == "__main__":
    unittest.main()
