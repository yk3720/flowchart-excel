"""スマート・パレット — Excel上に単発図形＋コネクタを直接生成する。

図形種別の判定は `layout_preview.shape_kind_for_type` / `SHAPE_CODE_BY_KIND`（SSOT）を再利用する。
コネクタの「浮いた終点」はアンカー図形を作らず、`AddConnector` の終点座標を直接指定して表現する
（アンカーを作って後で消す方式は削除漏れによる残留リスクがあるため採らない）。
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional

import pywintypes

from app.constants import ExcelConstants
from app.core.connector_manager import add_decision_label
from app.core.layout_preview import SHAPE_CODE_BY_KIND, shape_kind_for_type
from app.core.shape_placer import set_text_style

logger = logging.getLogger("flowchart-excel")


def create_smart_shape(
    sheet: Any,
    stype: str,
    left: float,
    top: float,
    width: float,
    height: float,
    theme: Dict[str, Any],
) -> Optional[str]:
    """図形＋コネクタ（＋判断ならYes/Noラベル）を生成し、1グループにまとめて返す。

    Returns:
        グループ名（Undo対象）。グループ化に失敗した場合は None。
    """
    kind, is_diamond = shape_kind_for_type(stype)
    stype_code = SHAPE_CODE_BY_KIND[kind]

    if is_diamond:
        height = height * 1.3
    elif kind == "oval":
        side = min(width, height)
        width = height = side

    shp = sheet.Shapes.AddShape(stype_code, left, top, width, height)
    shp.Fill.ForeColor.RGB = 0xFFFFFF
    shp.Line.ForeColor.RGB = theme["shape_line"]
    set_text_style(shp, "XXXX", is_manual=(kind == "manual"))

    group_names: List[str] = [shp.Name]

    try:
        # 下側コネクタ: 終点は「図形の下側から高さ分下」の座標に固定し、
        # BeginConnect のみ行う（EndConnect しない＝アンカー図形が不要）。
        conn_bottom = sheet.Shapes.AddConnector(
            ExcelConstants.MSOCONNECTOR_ELBOW,
            left + width / 2, top + height,
            left + width / 2, top + height * 2,
        )
        conn_bottom.ConnectorFormat.BeginConnect(shp, ExcelConstants.CONNECTOR_SITE_BOTTOM)
        conn_bottom.Line.ForeColor.RGB = theme["connector"]
        conn_bottom.Line.Weight = 2.25
        conn_bottom.Line.EndArrowheadStyle = 3
        group_names.append(conn_bottom.Name)

        if is_diamond:
            lbl_yes = add_decision_label(sheet, shp, "down")
            if lbl_yes:
                group_names.append(lbl_yes)
            else:
                logger.warning("decision_label_yes_creation_failed")

            conn_right = sheet.Shapes.AddConnector(
                ExcelConstants.MSOCONNECTOR_ELBOW,
                left + width, top + height / 2,
                left + width * 2, top + height / 2,
            )
            conn_right.ConnectorFormat.BeginConnect(shp, ExcelConstants.CONNECTOR_SITE_RIGHT)
            conn_right.Line.ForeColor.RGB = theme["connector"]
            conn_right.Line.Weight = 2.25
            conn_right.Line.EndArrowheadStyle = 3
            group_names.append(conn_right.Name)

            lbl_no = add_decision_label(sheet, shp, "right")
            if lbl_no:
                group_names.append(lbl_no)
            else:
                logger.warning("decision_label_no_creation_failed")
    except (pywintypes.com_error, AttributeError) as e:
        logger.error("smart_palette_connector_failed | error=%s", e)

    group_name: Optional[str] = None
    try:
        if len(group_names) > 1:
            group_name = sheet.Shapes.Range(tuple(group_names)).Group().Name
        else:
            group_name = group_names[0]
    except (pywintypes.com_error, AttributeError) as e:
        logger.error("smart_palette_grouping_failed | error=%s | members=%s", e, group_names)
        group_name = None

    logger.info(
        "smart_palette_shape_created | type=%s | position=(%s, %s) | size=(%s, %s) | group=%s",
        stype, left, top, width, height, group_name,
    )
    return group_name
