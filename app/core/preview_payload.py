"""Excel 表 → flowchart-studio プレビュー用 JSON ペイロード。"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from app.core.parse_table import _detect_ten_col_v2, parse_table_rows


def _cell_value(value: Any) -> Any:
    if value is None:
        return None
    return value


def excel_matrix_to_table(data: Tuple[Any, ...]) -> List[List[Any]]:
    """COM の Value タプルを studio `FlowTableRow[]` に変換する。"""
    rows: List[List[Any]] = []
    if not data:
        return rows
    if not isinstance(data[0], tuple):
        return [[_cell_value(data[0] if not isinstance(data, tuple) else data)]]
    for row in data:
        rows.append([_cell_value(c) for c in row])
    return rows


def table_list_to_com_tuple(table: List[List[Any]]) -> Tuple[Any, ...]:
    """studio プレビュー JSON の table を parse_table 互換タプルへ戻す。"""
    return tuple(tuple(row) for row in table)


def resolve_schema(table: List[List[Any]]) -> Optional[str]:
    if not table:
        return None
    col_count = len(table[0])
    as_tuple = tuple(tuple(r) for r in table)
    if col_count >= 10 and _detect_ten_col_v2(as_tuple, col_count):
        return "table-10col-v2"
    if col_count >= 10:
        return "table-10col-v1"
    if col_count >= 9:
        return "table-9col-v1"
    return None


def build_studio_preview_payload(
    data: Tuple[Any, ...],
    *,
    title: str,
    is_full_mode: bool,
    config: Dict[str, Any],
) -> Dict[str, Any]:
    """WebView に渡すペイロード（generateFlowchart 互換）。"""
    table = excel_matrix_to_table(data)
    nodes, _, col_count = parse_table_rows(data)
    if not nodes:
        raise ValueError("有効なノードがありません（ID が数値の行を確認してください）")

    schema = resolve_schema(table)
    return {
        "title": title,
        "isFullMode": bool(is_full_mode),
        "schema": schema,
        "table": table,
        "layout": {
            "width": float(config["width"]),
            "heightMin": float(config["height"]),
            "gapV": float(config["gap_v"]),
            "gapH": float(config["gap_h"]),
            "baseLeft": 40.0,
            "baseTop": 40.0,
        },
        "meta": {
            "nodeCount": len(nodes),
            "colCount": col_count,
        },
    }
