"""表データのパース — flowchart-studio table-10col-v2 互換。

SSOT: yk-application/flowchart-studio/lib/flowchart/table/parseTable.ts
"""
from __future__ import annotations

import logging
import re
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger("flowchart-excel")

TABLE_HEADERS_10_V2 = (
    "ID",
    "図形種別",
    "色",
    "接続先(下)",
    "接続先(右)",
    "段",
    "列",
    "Text1",
    "Text2",
    "Text3",
)

_SHAPE_ALIASES = {
    "開始": "端子",
    "終了": "端子",
    "データ": "入出力",
}


def _is_oval_type(text: str) -> bool:
    return any(token in text for token in ("〇", "○", "省略記号"))


def norm_id(value: Any) -> str:
    if value is None or value == "":
        return ""
    return str(value).split(".")[0].strip()


def parse_level(value: Any) -> int:
    if value is None or value == "":
        return 0
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def parse_level_optional(value: Any) -> Optional[int]:
    """parse_level 同様だが空欄・不正値を None で返す（C-2 の空欄判定専用）。

    parse_level は空欄を 0 に丸めるため、既存の描画・レイアウト経路
    （layout_preview.py・connector_manager.py）が依存するこの丸めは変更しない。
    C-2 の固定点反復は「空欄かどうか」を区別する必要があるため、生セル値を
    この関数経由で読む（parse_table_rows() が返す node["level"] は使わない）。
    """
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def split_dests(value: Any) -> List[str]:
    if value is None or value == "":
        return []
    return [norm_id(part) for part in str(value).split(",") if norm_id(part)]


def normalize_shape_type(raw: Any) -> str:
    text = str(raw).strip() if raw not in (None, "") else "処理"
    if "判断" in text:
        return "判断"
    if _is_oval_type(text):
        return "〇"
    for key, mapped in _SHAPE_ALIASES.items():
        if key in text:
            return mapped
    return text


_V2_DETECTION_SAMPLE_SIZE = 5


def _v2_signal_for_row(row: Any) -> bool:
    """1行だけを見た場合の v2（色が3列目）判定。"""
    third = row[2] if len(row) > 2 else None
    fourth = row[3] if len(row) > 3 else None
    if third is None or third == "":
        return True
    third_text = str(third).strip()
    if "," in third_text or third_text.isdigit():
        return False
    if fourth is not None and str(fourth).strip().replace(".", "", 1).isdigit():
        return True
    return True


def _detect_ten_col_v2(data: Tuple[Any, ...], col_count: int) -> bool:
    """10列 v2（色が3列目）か v1（接続先が3列目）かを推定する。

    1行のみのヒューリスティックはその行の内容（例: 判断ノードで下方向未使用）次第で
    テーブル全体を誤判定しうるため、先頭から最大 `_V2_DETECTION_SAMPLE_SIZE` 行の
    有効ID行を集めて多数決にする。
    """
    if col_count < 10:
        return False
    sample = [row for row in data if norm_id(row[0]).isdigit()][:_V2_DETECTION_SAMPLE_SIZE]
    if not sample:
        return True
    votes = [_v2_signal_for_row(row) for row in sample]
    true_count = sum(1 for v in votes if v)
    false_count = len(votes) - true_count
    return true_count >= false_count


def parse_table_rows(
    data: Tuple[Any, ...],
    *,
    force_v2: Optional[bool] = None,
) -> Tuple[List[Dict[str, Any]], Dict[int, List[Dict[str, Any]]], int]:
    """Excel セル配列をノードリストと行マップに変換する。

    Returns:
        nodes, row_map, col_count
    """
    if not data:
        return [], {}, 0

    col_count = len(data[0]) if data[0] else 0
    is_v2 = force_v2 if force_v2 is not None else _detect_ten_col_v2(data, col_count)

    nodes: List[Dict[str, Any]] = []
    row_map: Dict[int, List[Dict[str, Any]]] = {}
    seen_ids: set[str] = set()

    for i, row in enumerate(data):
        row_list = list(row) if isinstance(row, tuple) else row
        nid = norm_id(row_list[0] if row_list else "")
        if not nid or not re.fullmatch(r"\d+", nid):
            continue

        type_raw = row_list[1] if len(row_list) > 1 else None
        if type_raw is None or str(type_raw).strip() == "":
            continue  # ID はあるが図形種別が空欄 → フローから除外（描画・接続先解決のいずれの対象にもしない）

        if nid in seen_ids:
            logger.warning("duplicate_id_skipped | id=%s | row=%s", nid, i)
            continue  # 仕様§4.2: 最初に見つかったノードのみ使用（後続は無視）
        seen_ids.add(nid)

        txts: List[str]
        dests_down: List[str]
        dests_right: List[str]
        level: int
        tier: Optional[int] = None
        color_hint: Optional[str] = None

        if col_count >= 9 and is_v2:
            txts = [
                str(row_list[j])
                for j in range(7, min(10, len(row_list)))
                if row_list[j] not in (None, "")
            ]
            color_raw = row_list[2] if len(row_list) > 2 else None
            if color_raw not in (None, ""):
                color_hint = str(color_raw).strip()
            dests_down = split_dests(row_list[3] if len(row_list) > 3 else None)
            dests_right = split_dests(row_list[4] if len(row_list) > 4 else None)
            tier = parse_level(row_list[5] if len(row_list) > 5 else None)
            level = parse_level(row_list[6] if len(row_list) > 6 else None)
        elif col_count >= 9:
            txts = [
                str(row_list[j])
                for j in range(6, min(9, len(row_list)))
                if row_list[j] not in (None, "")
            ]
            dests_down = split_dests(row_list[2] if len(row_list) > 2 else None)
            dests_right = split_dests(row_list[3] if len(row_list) > 3 else None)
            tier = parse_level(row_list[4] if len(row_list) > 4 else None)
            level = parse_level(row_list[5] if len(row_list) > 5 else None)
            if col_count >= 10 and row_list[9] not in (None, ""):
                color_hint = str(row_list[9]).strip()
        elif col_count >= 8:
            txts = [
                str(row_list[j])
                for j in range(5, min(8, len(row_list)))
                if row_list[j] not in (None, "")
            ]
            dests_down = split_dests(row_list[2] if len(row_list) > 2 else None)
            dests_right = split_dests(row_list[3] if len(row_list) > 3 else None)
            level = parse_level(row_list[4] if len(row_list) > 4 else None)
        elif col_count == 7:
            txts = [
                str(row_list[j])
                for j in range(4, min(7, len(row_list)))
                if row_list[j] not in (None, "")
            ]
            dests_down = split_dests(row_list[2] if len(row_list) > 2 else None)
            dests_right = []
            level = parse_level(row_list[3] if len(row_list) > 3 else None)
        else:
            txts = (
                [str(row_list[2])]
                if len(row_list) > 2 and row_list[2] not in (None, "")
                else []
            )
            dests_down = split_dests(row_list[3] if len(row_list) > 3 else None)
            dests_right = []
            level = parse_level(row_list[4] if len(row_list) > 4 else 0)

        node: Dict[str, Any] = {
            "id": nid,
            "type": normalize_shape_type(type_raw),
            "full_text": "\n".join(txts),
            "dests_down": dests_down,
            "dests_right": dests_right,
            "level": level,
            "ridx": i,
        }
        if tier is not None:
            node["tier"] = tier
        if color_hint:
            node["color_hint"] = color_hint

        nodes.append(node)
        row_map.setdefault(i, []).append(node)

    return nodes, row_map, col_count
