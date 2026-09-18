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
    frame_anchor_offset,
)
from app.core.layout_preview import PreviewModel, build_preview_model, estimate_row_heights
from app.core.parse_table import parse_table_rows
from app.core.preview_payload import (
    build_studio_preview_payload,
    schema_to_force_v2,
    table_list_to_com_tuple,
)
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

    def build_studio_payload(
        self,
        is_full_mode: bool,
        config: Dict[str, Any],
    ) -> Dict[str, Any]:
        """flowchart-studio 同等プレビュー用 JSON ペイロード。"""
        data, sheet, start_cell, title_txt = self._read_selection(is_full_mode)
        sel = get_excel_app().Selection
        r_tgt = sel.CurrentRegion if is_full_mode else sel
        payload = build_studio_preview_payload(
            data,
            title=title_txt,
            is_full_mode=is_full_mode,
            config=config,
        )
        meta = dict(payload.get("meta") or {})
        meta["live"] = True
        meta["watch"] = {
            "isFullMode": bool(is_full_mode),
            "workbookName": str(sheet.Parent.Name),
            "sheetName": str(sheet.Name),
            "anchorAddress": str(start_cell.Address),
            "rangeAddress": str(r_tgt.Address),
        }
        payload["meta"] = meta
        return payload

    def _read_current_anchor(self) -> Tuple[Any, Any]:
        """「Excelに作成」実行時点で選択されているセルを描画起点として取得する。

        表の読み込み元とは独立に、作成ボタンを押した瞬間の選択セルを起点にする
        （表の位置を覚えておく方式は行挿入等でズレるため採らない）。
        """
        app = get_excel_app()
        if not app:
            raise RuntimeError("Excelが起動していません。")
        try:
            sheet = app.ActiveSheet
            start_cell = app.Selection.Cells(1, 1)
        except (pywintypes.com_error, AttributeError) as exc:
            raise RuntimeError(
                "作成先のセルを取得できませんでした。Excelでセルを選択してから再試行してください。"
            ) from exc
        try:
            logger.info(
                "draw_anchor_resolved | sheet=%s | address=%s | left=%s | top=%s",
                sheet.Name, start_cell.Address, start_cell.Left, start_cell.Top,
            )
        except (pywintypes.com_error, AttributeError):
            pass
        return sheet, start_cell

    def build_preview_from_payload(self, payload: Dict[str, Any]) -> PreviewModel:
        """プレビュー用ペイロード（Excel再読込なし）からプレビュー用モデルを返す。

        Canvasフォールバック専用。`build_studio_payload` で一度読んだ `payload` を
        そのまま使うことで、埋め込み/2窓WebViewと同じ「表示＝作成」契約を満たす。
        """
        table = payload.get("table") or []
        data = table_list_to_com_tuple(table)
        layout = payload.get("layout") or {}
        force_v2 = schema_to_force_v2(payload.get("schema"))
        nodes, row_map, col_count = parse_table_rows(data, force_v2=force_v2)
        title_txt = str(payload.get("title") or "フローチャート")
        is_full_mode = bool(payload.get("isFullMode"))
        config = {
            "width": float(layout["width"]),
            "height": float(layout["heightMin"]),
            "gap_v": float(layout["gapV"]),
            "gap_h": float(layout["gapH"]),
        }
        logger.info(
            "preview_from_payload_parse | nodes=%s | col_count=%s | full=%s",
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

    def draw_from_studio_payload(
        self,
        payload: Dict[str, Any],
        theme: Dict[str, Any],
    ) -> str:
        """プレビュー確定スナップショットから描画（内容＝表示スナップショット、位置＝実行時点の選択セル）。"""
        table = payload.get("table") or []
        if not table:
            raise ValueError("スナップショットに表データがありません。")

        sheet, start_cell = self._read_current_anchor()
        data = table_list_to_com_tuple(table)
        layout = payload.get("layout") or {}
        config = {
            "width": float(layout["width"]),
            "height": float(layout["heightMin"]),
            "gap_v": float(layout["gapV"]),
            "gap_h": float(layout["gapH"]),
        }
        title_txt = str(payload.get("title") or "フローチャート")
        is_full_mode = bool(payload.get("isFullMode"))
        force_v2 = schema_to_force_v2(payload.get("schema"))

        logger.info(
            "draw_from_snapshot | title=%s | full=%s | rows=%s",
            title_txt,
            is_full_mode,
            len(table),
        )
        return self._draw_core(
            data=data,
            sheet=sheet,
            start_cell=start_cell,
            title_txt=title_txt,
            is_full_mode=is_full_mode,
            config=config,
            theme=theme,
            force_v2=force_v2,
        )

    def _draw_core(
        self,
        *,
        data: Any,
        sheet: Any,
        start_cell: Any,
        title_txt: str,
        is_full_mode: bool,
        config: Dict[str, Any],
        theme: Dict[str, Any],
        force_v2: Optional[bool] = None,
    ) -> str:
        app = get_excel_app()
        if not app:
            logger.error("excel_not_found | Excel is not running.")
            raise RuntimeError("Excelが起動していません。")

        app.ScreenUpdating = False
        app.DisplayAlerts = False
        created_names: List[str] = []

        def _rollback() -> None:
            for name in created_names:
                try:
                    sheet.Shapes(name).Delete()
                except (pywintypes.com_error, AttributeError):
                    pass

        try:
            base_left = float(start_cell.Left)
            base_top = float(start_cell.Top)
            if is_full_mode:
                # 外枠+タイトルを付ける場合、選択セル＝外枠の角になるよう
                # 図形群の原点をその分だけ右下へずらす（add_frame_and_title 参照）
                dx, dy = frame_anchor_offset()
                base_left += dx
                base_top += dy
            h_min = float(config["height"])
            w_fix = float(config["width"])
            gv = float(config["gap_v"])
            gh = float(config["gap_h"])

            nodes, row_map, col_count = parse_table_rows(data, force_v2=force_v2)
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
            created_names.extend(standalone_names)
            created_names.extend(info["shp"].Name for info in diamond_info)

            if self.stop_event.is_set():
                logger.info("draw_cancelled_before_connect")
                _rollback()
                return ""

            connector_names = connect_nodes(
                sheet, nodes, shape_map, theme, self.stop_event
            )
            created_names.extend(connector_names)

            if self.stop_event.is_set():
                logger.info("draw_cancelled_before_finalize")
                _rollback()
                return ""

            composite_pairs = finalize_composites(sheet, diamond_info, w_fix)
            created_names.extend(tx.Name for _, tx in composite_pairs)

            if self.stop_event.is_set():
                logger.info("draw_cancelled_before_groups")
                _rollback()
                return ""

            extra_names: List[str] = []
            if is_full_mode:
                extra_names = add_frame_and_title(sheet, bounds, title_txt)
                created_names.extend(extra_names)

            all_names = standalone_names + connector_names + extra_names
            group_name = create_final_groups(sheet, all_names, composite_pairs)

            logger.info("draw_completed | group_name=%s", group_name)
            return group_name

        except Exception:
            logger.exception("draw_failed_rolling_back")
            _rollback()
            raise
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
