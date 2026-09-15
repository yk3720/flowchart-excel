"""フロー配置の dry-run（Excel COM なし）。

place_shapes と同じ tier/level 規則で幾何だけ計算する。
SSOT 配置: flowchart-studio layoutGrid.ts · app/core/shape_placer.py
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Literal, Optional, Tuple

from app.constants import ExcelConstants

ShapeKind = Literal["rect", "diamond", "roundrect", "parallelogram", "manual", "oval"]

# 図形種別(ShapeKind) → Excel AutoShape の MSOコード。SSOT: 本辞書のみ
# (旧 shape_placer._SHAPE_CODE / main_window._smart_input の自前 if/elif を統合)。
SHAPE_CODE_BY_KIND: Dict[ShapeKind, int] = {
    "rect": ExcelConstants.MSOSHAPE_RECTANGLE,
    "diamond": ExcelConstants.MSOSHAPE_DIAMOND,
    "roundrect": ExcelConstants.MSOSHAPE_ROUNDED_RECTANGLE,
    "parallelogram": ExcelConstants.MSOSHAPE_PARALLELOGRAM,
    "manual": ExcelConstants.MSOSHAPE_MANUAL_INPUT,
    "oval": ExcelConstants.MSOSHAPE_OVAL,
}


@dataclass
class PlacedNode:
    """配置済みノード（プレビュー・描画共通）。"""

    id: str
    type: str
    full_text: str
    level: int
    tier: int
    left: float
    top: float
    width: float
    height: float
    shape_kind: ShapeKind
    is_diamond: bool
    dests_down: List[str] = field(default_factory=list)
    dests_right: List[str] = field(default_factory=list)
    color_hint: Optional[str] = None


@dataclass
class PreviewEdge:
    source_id: str
    target_id: str
    direction: Literal["down", "right"]
    is_decision: bool


@dataclass
class PreviewModel:
    title: str
    is_full_mode: bool
    nodes: List[PlacedNode]
    edges: List[PreviewEdge]
    bounds: Tuple[float, float, float, float]
    node_count: int
    edge_count: int
    warnings: List[str] = field(default_factory=list)


def _node_tier(node: Dict[str, Any]) -> int:
    tier = node.get("tier")
    if tier is not None:
        return int(tier)
    return int(node["ridx"])


def shape_kind_for_type(stype: str) -> Tuple[ShapeKind, bool]:
    """表の図形種別文字列 → (ShapeKind, is_diamond)。図形種別マッピングのSSOT。"""
    if "判断" in stype:
        return "diamond", True
    if any(x in stype for x in ("〇", "○", "省略記号")):
        return "oval", False
    if any(x in stype for x in ["端子", "開始", "終了"]):
        return "roundrect", False
    if any(x in stype for x in ["入出力", "データ"]):
        return "parallelogram", False
    if "手動入力" in stype:
        return "manual", False
    return "rect", False


def estimate_row_heights(
    row_map: Dict[int, List[Dict[str, Any]]],
    h_min: float,
) -> Dict[int, float]:
    """テキスト行数から高さ概算（Excel AutoSize の代替）。"""
    heights: Dict[int, float] = {}
    for ri, row_nodes in row_map.items():
        max_h = h_min
        for node in row_nodes:
            lines = max(1, str(node.get("full_text") or "").count("\n") + 1)
            max_h = max(max_h, float(lines) * 18.0 + 15.0)
        heights[ri] = max_h
    return heights


def compute_layout(
    row_map: Dict[int, List[Dict[str, Any]]],
    row_heights: Dict[int, float],
    base_left: float,
    base_top: float,
    w_fix: float,
    gv: float,
    gh: float,
    h_min: float,
) -> Tuple[List[PlacedNode], Tuple[float, float, float, float]]:
    """tier/level でノード矩形を計算する（COM なし）。"""
    placed: List[PlacedNode] = []
    l_list: List[float] = []
    t_list: List[float] = []
    r_list: List[float] = []
    b_list: List[float] = []

    tier_map: Dict[int, Dict[str, Any]] = {}
    for ri in sorted(row_map.keys()):
        for node in row_map[ri]:
            tier = _node_tier(node)
            bucket = tier_map.setdefault(
                tier,
                {"tier": tier, "nodes": [], "height": h_min},
            )
            bucket["nodes"].append(node)
            bucket["height"] = max(bucket["height"], row_heights.get(ri, h_min))

    current_top = base_top
    last_tier: Optional[int] = None

    for tier in sorted(tier_map.keys()):
        bucket = tier_map[tier]
        if last_tier is not None:
            prev = tier_map[last_tier]
            current_top += prev["height"] + gv

        for node in sorted(bucket["nodes"], key=lambda n: (n["level"], n["id"])):
            cell_left = base_left + node["level"] * (w_fix + gh)
            kind, is_diamond = shape_kind_for_type(str(node["type"]))
            row_h = float(bucket["height"])
            if kind == "oval":
                shp_w = min(w_fix, row_h)
                shp_h = shp_w
                left_pos = cell_left + (w_fix - shp_w) / 2
                top_pos = current_top
            else:
                shp_w = w_fix
                shp_h = row_h * 1.3 if is_diamond else row_h
                v_off = (shp_h - row_h) / 2 if is_diamond else 0.0
                left_pos = cell_left
                top_pos = current_top - v_off

            color_raw = node.get("color_hint")
            color_hint = str(color_raw).strip() if color_raw not in (None, "") else None

            placed.append(
                PlacedNode(
                    id=str(node["id"]),
                    type=str(node["type"]),
                    full_text=str(node.get("full_text") or ""),
                    level=int(node["level"]),
                    tier=tier,
                    left=left_pos,
                    top=top_pos,
                    width=shp_w,
                    height=shp_h,
                    shape_kind=kind,
                    is_diamond=is_diamond,
                    dests_down=list(node.get("dests_down") or []),
                    dests_right=list(node.get("dests_right") or []),
                    color_hint=color_hint,
                )
            )
            l_list.append(left_pos)
            t_list.append(top_pos)
            r_list.append(left_pos + shp_w)
            b_list.append(top_pos + shp_h)

        last_tier = tier

    bounds = (
        (min(l_list), min(t_list), max(r_list), max(b_list)) if l_list else (0.0, 0.0, 0.0, 0.0)
    )
    return placed, bounds


def build_edges(nodes: List[PlacedNode]) -> Tuple[List[PreviewEdge], List[str]]:
    """接続先からエッジを組み立て、欠落 ID を警告する。"""
    by_id = {n.id: n for n in nodes}
    edges: List[PreviewEdge] = []
    warnings: List[str] = []
    missing: set[str] = set()

    for node in nodes:
        is_decision = "判断" in node.type
        for direction, dests in (
            ("down", node.dests_down),
            ("right", node.dests_right),
        ):
            for dest_id in dests:
                if dest_id not in by_id:
                    missing.add(f"{node.id}→{dest_id}（{direction}）")
                    continue
                edges.append(
                    PreviewEdge(
                        source_id=node.id,
                        target_id=dest_id,
                        direction=direction,  # type: ignore[arg-type]
                        is_decision=is_decision,
                    )
                )

    if missing:
        warnings.append("接続先が見つかりません: " + ", ".join(sorted(missing)))
    return edges, warnings


def build_preview_model(
    *,
    nodes_raw: List[Dict[str, Any]],
    row_map: Dict[int, List[Dict[str, Any]]],
    config: Dict[str, Any],
    title: str,
    is_full_mode: bool,
    base_left: float = 0.0,
    base_top: float = 0.0,
    row_heights: Optional[Dict[int, float]] = None,
) -> PreviewModel:
    """パース結果からプレビューモデルを構築する。"""
    h_min = float(config["height"])
    w_fix = float(config["width"])
    gv = float(config["gap_v"])
    gh = float(config["gap_h"])

    heights = row_heights if row_heights is not None else estimate_row_heights(row_map, h_min)
    placed, bounds = compute_layout(
        row_map, heights, base_left, base_top, w_fix, gv, gh, h_min
    )
    edges, warnings = build_edges(placed)
    if not nodes_raw:
        warnings.append("有効なノードがありません（ID が数値の行を確認してください）")

    return PreviewModel(
        title=title,
        is_full_mode=is_full_mode,
        nodes=placed,
        edges=edges,
        bounds=bounds,
        node_count=len(placed),
        edge_count=len(edges),
        warnings=warnings,
    )
