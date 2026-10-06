"""段(tier) の提案（C-3 + 空欄補完）。

Design doc: docs/04_レビュー/構想設計_F3_段列自動計算提案機能_2026-09-18.md

- blank_only: 段が空欄の行へ、表の並び順で連番を割り当てる（既存の最大段+1 から）。
  全部空欄のときは 0,1,2… とし、段・列未設定による座標重複を解消する。
- full_recalc: 既存の段のユニーク値を昇順ランクへ振り直す（C-3）。同じ段を共有する
  ノード集合は維持され、レイアウトの見た目は変わらない。
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence

from app.core.level_inference import LevelProposal, ProposalMode, ProposalResult


def compute_tier_proposals(
    nodes: Sequence[Dict[str, Any]],
    raw_tiers: Dict[str, Optional[int]],
    *,
    mode: ProposalMode = "blank_only",
) -> ProposalResult:
    ordered = sorted(nodes, key=lambda n: int(n.get("ridx", 0)))
    if not ordered:
        return ProposalResult()

    if mode == "blank_only":
        return _blank_only(ordered, raw_tiers)
    return _full_recalc_ranks(ordered, raw_tiers)


def _blank_only(
    ordered: Sequence[Dict[str, Any]],
    raw_tiers: Dict[str, Optional[int]],
) -> ProposalResult:
    existing = [
        raw_tiers[n["id"]]
        for n in ordered
        if raw_tiers.get(n["id"]) is not None
    ]
    next_tier = (max(existing) + 1) if existing else 0
    proposals: List[LevelProposal] = []
    for node in ordered:
        nid = node["id"]
        current = raw_tiers.get(nid)
        if current is not None:
            continue
        proposals.append(
            LevelProposal(
                node_id=nid,
                current=None,
                proposed=next_tier,
                reason="表の並び順（空欄の段）",
            )
        )
        next_tier += 1
    return ProposalResult(proposals=proposals)


def _full_recalc_ranks(
    ordered: Sequence[Dict[str, Any]],
    raw_tiers: Dict[str, Optional[int]],
) -> ProposalResult:
    """ユニークな段値を昇順ソートしたランク（0,1,2…）へ振り直す。"""
    values: Dict[str, int] = {}
    for node in ordered:
        nid = node["id"]
        raw = raw_tiers.get(nid)
        values[nid] = 0 if raw is None else int(raw)

    unique = sorted(set(values.values()))
    rank = {value: index for index, value in enumerate(unique)}

    proposals: List[LevelProposal] = []
    for node in ordered:
        nid = node["id"]
        current_raw = raw_tiers.get(nid)
        current_value = values[nid]
        proposed = rank[current_value]
        if current_value == proposed and current_raw is not None:
            continue
        proposals.append(
            LevelProposal(
                node_id=nid,
                current=current_raw,
                proposed=proposed,
                reason="段の値をユニーク順位で振り直し",
            )
        )
    return ProposalResult(proposals=proposals)
