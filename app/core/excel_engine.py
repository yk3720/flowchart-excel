"""Excel 操作・描画エンジン — 10列表駆動。"""
import logging
import threading
from typing import Any, Dict, List, Optional, Tuple

import pywintypes
import win32com.client

from app.constants import ExcelConstants
from app.core.connector_manager import connect_nodes
from app.core.group_manager import (
    add_frame_and_title,
    create_final_groups,
    finalize_composites,
)
from app.core.layout_preview import PreviewModel, build_preview_model, estimate_row_heights
from app.core.parse_table import parse_table_rows
from app.core.shape_placer import place_shapes

logger = logging.getLogger("flowchart-excel")


def get_excel_app() -> Optional[Any]:
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except (pywintypes.com_error, AttributeError):
        return None


class ExcelFlowchartEngine:
    """Excel 上でのフローチャート描画エンジン。"""

    def __init__(self, stop_event: threading.Event) -> None:
        self.stop_event = stop_event
        self.last_group_name: Optional[str] = None

    def _read_selection(
        self, is_full_mode: bool
    ) -> Tuple[Any, Any, Any, str]:
        """選択範囲のセル配列・シート・起点・タイトルを取得する。"""
        app = get_excel_app()
        if not app:
            logger.error("excel_not_found | Excel is not running.")
            raise RuntimeError("Excelが起動していません。")

        sel = app.Selection
        r_tgt = sel.CurrentRegion if is_full_mode else sel
        data = r_tgt.Value

        if not data or not isinstance(data, tuple):
            logger.warning("no_data_selected | Selection is empty or invalid.")
            raise ValueError("選択範囲にデータがありません。")

        sheet = app.ActiveSheet
        start_cell = r_tgt.Cells(1, 1)

        title_txt = "フローチャート"
        if is_full_mode:
            for i in range(-5, 1):
                row_idx = max(1, start_cell.Row + i)
                cell = sheet.Cells(row_idx, start_cell.Column)
                if cell.Interior.Color == ExcelConstants.TITLE_BG_COLOR and cell.Value:
                    title_txt = str(cell.Value)
                    break

        return data, sheet, start_cell, title_txt

    def build_preview(
        self,
        is_full_mode: bool,
        config: Dict[str, Any],
    ) -> PreviewModel:
        """Excel を読まず描画せず、プレビュー用モデルを返す。"""
        data, _sheet, _start_cell, title_txt = self._read_selection(is_full_mode)
        nodes, row_map, col_count = parse_table_rows(data)
        logger.info(
            "preview_parse | nodes=%s | col_count=%s | full=%s",
            len(nodes),
            col_count,
            is_full_mode,
        )
        return build_preview_model(
            nodes_raw=nodes,
            row_map=row_map,
            config=config,
            title=title_txt,
            is_full_mode=is_full_mode,
            row_heights=estimate_row_heights(row_map, float(config["height"])),
        )

    def draw(
        self,
        is_full_mode: bool,
        config: Dict[str, Any],
        theme: Dict[str, Any],
    ) -> str:
        app = get_excel_app()
        if not app:
            logger.error("excel_not_found | Excel is not running.")
            raise RuntimeError("Excelが起動していません。")

        app.ScreenUpdating = False
        app.DisplayAlerts = False

        try:
            data, sheet, start_cell, title_txt = self._read_selection(is_full_mode)

            base_left = float(start_cell.Left)
            base_top = float(start_cell.Top)
            h_min = float(config["height"])
            w_fix = float(config["width"])
            gv = float(config["gap_v"])
            gh = float(config["gap_h"])

            nodes, row_map, col_count = parse_table_rows(data)
            if not nodes:
                return ""

            logger.info(
                "parse_completed | nodes=%s | col_count=%s | has_tier=%s",
                len(nodes),
                col_count,
                any("tier" in n for n in nodes),
            )

            row_heights = self._calculate_row_heights(sheet, row_map, w_fix, h_min)

            shape_map, standalone_names, diamond_info, bounds = place_shapes(
                sheet,
                row_map,
                row_heights,
                base_left,
                base_top,
                w_fix,
                gv,
                gh,
                theme,
                self.stop_event,
                h_min,
            )

            if self.stop_event.is_set():
                logger.info("draw_cancelled_before_connect")
                return ""

            connector_names = connect_nodes(
                sheet, nodes, shape_map, theme, self.stop_event
            )

            if self.stop_event.is_set():
                logger.info("draw_cancelled_before_finalize")
                return ""

            composite_pairs = finalize_composites(sheet, diamond_info, w_fix)

            extra_names: List[str] = []
            if is_full_mode:
                extra_names = add_frame_and_title(sheet, bounds, title_txt)

            all_names = standalone_names + connector_names + extra_names
            group_name = create_final_groups(sheet, all_names, composite_pairs)

            logger.info("draw_completed | group_name=%s", group_name)
            return group_name

        finally:
            app.ScreenUpdating = True
            app.DisplayAlerts = True

    def _calculate_row_heights(
        self,
        sheet: Any,
        row_map: Dict[int, List[Dict[str, Any]]],
        w_fix: float,
        h_min: float,
    ) -> Dict[int, float]:
        heights: Dict[int, float] = {}
        temp_shp = sheet.Shapes.AddShape(
            ExcelConstants.MSOSHAPE_RECTANGLE, -5000, -5000, w_fix, h_min
        )
        try:
            for ri, row_nodes in row_map.items():
                max_h = h_min
                for node in row_nodes:
                    temp_shp.TextFrame2.TextRange.Text = node["full_text"]
                    temp_shp.TextFrame2.AutoSize = 1
                    max_h = max(max_h, float(temp_shp.Height) + 15.0)
                heights[ri] = max_h
        finally:
            temp_shp.Delete()
        return heights
