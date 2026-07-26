"""図形配置 — 10列（段+列）レイアウト対応。

SSOT: yk-application/flowchart-studio/lib/flowchart/graph/layoutGrid.ts
配置幾何: app/core/layout_preview.compute_layout
"""
import logging
from typing import Any, Dict, List, Tuple

import pywintypes

from app.constants import ExcelConstants, FONT_FAMILY
from app.core.layout_preview import compute_layout

logger = logging.getLogger("flowchart-excel")

_SHAPE_CODE = {
    "rect": ExcelConstants.MSOSHAPE_RECTANGLE,
    "diamond": ExcelConstants.MSOSHAPE_DIAMOND,
    "roundrect": ExcelConstants.MSOSHAPE_ROUNDED_RECTANGLE,
    "parallelogram": ExcelConstants.MSOSHAPE_PARALLELOGRAM,
    "manual": ExcelConstants.MSOSHAPE_MANUAL_INPUT,
}


def place_shapes(
    sheet: Any,
    row_map: Dict[int, List[Dict[str, Any]]],
    row_heights: Dict[int, float],
    base_left: float,
    base_top: float,
    w_fix: float,
    gv: float,
    gh: float,
    theme: Dict[str, Any],
    stop_event: Any,
    h_min: float,
) -> Tuple[Any, ...]:
    """図形を tier（段）ベースで配置する。"""
    placed, bounds = compute_layout(
        row_map, row_heights, base_left, base_top, w_fix, gv, gh, h_min
    )

    shape_map: Dict[str, Any] = {}
    standalone_names: List[str] = []
    diamond_info: List[Dict[str, Any]] = []

    for item in placed:
        if stop_event.is_set():
            break

        stype_code = _SHAPE_CODE[item.shape_kind]
        shp = sheet.Shapes.AddShape(
            stype_code,
            item.left,
            item.top,
            item.width,
            item.height,
        )
        shp.Fill.ForeColor.RGB = 0xFFFFFF
        shp.Line.ForeColor.RGB = theme["shape_line"]

        if item.is_diamond:
            # 菱形は後段でテキスト枠と合成するため、ここでは幾何のみ
            row_h = item.height / 1.3
            diamond_info.append(
                {
                    "shp": shp,
                    "l": item.left,
                    "t": item.top + (item.height - row_h) / 2,
                    "h": row_h,
                    "txt": item.full_text,
                }
            )
        else:
            standalone_names.append(shp.Name)
            set_text_style(
                shp,
                item.full_text,
                is_manual=(stype_code == ExcelConstants.MSOSHAPE_MANUAL_INPUT),
            )

        shape_map[item.id] = shp

    return shape_map, standalone_names, diamond_info, bounds


def set_text_style(shp: Any, text: str, is_manual: bool = False) -> None:
    """図形のテキストスタイルを一括設定する。"""
    try:
        tf2 = shp.TextFrame2
        tf2.TextRange.Text = text
        tf2.TextRange.ParagraphFormat.Alignment = ExcelConstants.MSO_ALIGN_CENTER
        tf2.VerticalAnchor = ExcelConstants.MSO_ANCHOR_MIDDLE
        tf2.WordWrap = True
        tf2.TextRange.Font.Size = 11
        tf2.TextRange.Font.Name = FONT_FAMILY
        tf2.TextRange.Font.Fill.ForeColor.RGB = 0
        if is_manual:
            tf2.TextRange.Font.Italic = True
    except (pywintypes.com_error, AttributeError) as exc:
        logger.warning("set_text_style_failed | error=%s", exc)
