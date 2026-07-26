"""コネクタ接続 — tier / level 対応。"""
import logging
from typing import Any, Dict, List, Optional

import pywintypes

from app.constants import ExcelConstants, FONT_FAMILY

logger = logging.getLogger("flowchart-excel")


def _node_tier(node: Dict[str, Any]) -> int:
    tier = node.get("tier")
    if tier is not None:
        return int(tier)
    return int(node["ridx"])


def connect_nodes(
    sheet: Any,
    nodes: List[Dict[str, Any]],
    shape_map: Dict[str, Any],
    theme: Dict[str, Any],
    stop_event: Any,
) -> List[str]:
    names: List[str] = []
    node_dict = {n["id"]: n for n in nodes}

    for node in nodes:
        source_shape = shape_map.get(node["id"])
        if not source_shape or stop_event.is_set():
            continue

        for direction, dests in [
            ("down", node["dests_down"]),
            ("right", node["dests_right"]),
        ]:
            for dest_id in dests:
                target_shape = shape_map.get(dest_id)
                target_node = node_dict.get(dest_id)
                if not target_shape or not target_node:
                    continue

                is_loop = target_node["ridx"] < node["ridx"] or _node_tier(
                    target_node
                ) < _node_tier(node)
                level_diff = target_node["level"] - node["level"]
                tier_diff = _node_tier(target_node) - _node_tier(node)

                source_site = ExcelConstants.CONNECTOR_SITE_BOTTOM
                target_site = ExcelConstants.CONNECTOR_SITE_TOP
                connector_type = ExcelConstants.MSOCONNECTOR_STRAIGHT

                if direction == "down":
                    if level_diff != 0 or is_loop or tier_diff != 1:
                        connector_type = ExcelConstants.MSOCONNECTOR_ELBOW
                        if level_diff < 0:
                            target_site = ExcelConstants.CONNECTOR_SITE_LEFT
                        elif level_diff > 0:
                            source_site = ExcelConstants.CONNECTOR_SITE_RIGHT
                            target_site = ExcelConstants.CONNECTOR_SITE_TOP
                else:
                    source_site = ExcelConstants.CONNECTOR_SITE_RIGHT
                    target_site = ExcelConstants.CONNECTOR_SITE_TOP
                    connector_type = ExcelConstants.MSOCONNECTOR_ELBOW
                    if level_diff == 0 and is_loop:
                        target_site = ExcelConstants.CONNECTOR_SITE_RIGHT
                    elif level_diff < 0:
                        target_site = ExcelConstants.CONNECTOR_SITE_LEFT

                if (
                    direction == "down"
                    and abs(float(source_shape.Left) - float(target_shape.Left)) < 5
                    and not is_loop
                    and tier_diff == 1
                ):
                    connector_type = ExcelConstants.MSOCONNECTOR_STRAIGHT

                try:
                    conn = sheet.Shapes.AddConnector(connector_type, 0, 0, 10, 10)
                    names.append(conn.Name)
                    conn.ConnectorFormat.BeginConnect(source_shape, source_site)
                    conn.ConnectorFormat.EndConnect(target_shape, target_site)
                    conn.Line.ForeColor.RGB = theme["connector"]
                    conn.Line.Weight = 2.25
                    conn.Line.EndArrowheadStyle = 3

                    if "判断" in node["type"]:
                        label_name = add_decision_label(sheet, source_shape, direction)
                        if label_name:
                            names.append(label_name)
                except (pywintypes.com_error, AttributeError) as exc:
                    logger.error(
                        "connector_error | from=%s | to=%s | error=%s",
                        node["id"],
                        dest_id,
                        exc,
                    )
    return names


def add_decision_label(sheet: Any, shape: Any, direction: str) -> Optional[str]:
    try:
        label_w, label_h = 30, 16
        if direction == "right":
            lx = shape.Left + shape.Width - 10
            ly = shape.Top + (shape.Height / 2) - label_h - 3
            text = "No"
        else:
            lx = shape.Left + (shape.Width / 2) - label_w - 3
            ly = shape.Top + shape.Height + 3
            text = "Yes"

        label = sheet.Shapes.AddTextbox(1, lx, ly, label_w, label_h)
        label.Fill.Visible = False
        label.Line.Visible = False
        text_range = label.TextFrame2.TextRange
        text_range.Text = text
        text_range.Font.Size = 9
        text_range.Font.Bold = True
        text_range.Font.Name = FONT_FAMILY
        text_range.Font.Fill.ForeColor.RGB = 0
        text_range.ParagraphFormat.Alignment = ExcelConstants.MSO_ALIGN_CENTER
        label.TextFrame2.VerticalAnchor = ExcelConstants.MSO_ANCHOR_MIDDLE
        return label.Name
    except (pywintypes.com_error, AttributeError):
        logger.warning("decision_label_creation_failed | direction=%s", direction)
        return None
