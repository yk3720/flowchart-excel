"""列(level) の空欄セルを隣接行から推測する固定点反復ロジック（C-2）。

Design doc: docs/04_レビュー/構想設計_F3_段列自動計算提案機能_2026-09-18.md

推測ルール（方針A）:
- 行Xが行Yの「接続先(下)」に指定されている → Xの列 = Yの列（継承）
- 行Xが行Yの「接続先(右)」に指定されている → Xの列 = Yの列＋1
- 複数の行から矛盾する推測が出た場合は自動確定させず要確認へ
- カンマ区切りの複数宛先を持つ行は推測の対象・根拠のいずれからも除外（MVPスコープ外）
- 循環（判断ノードの No 分岐による後方ループ）を含む実データに対応するため、
  再帰は使わず固定点反復（これ以上変化しなくなるまで1パスずつ繰り返す）で実装する。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Literal, Optional, Sequence, Tuple

ProposalMode = Literal["blank_only", "full_recalc"]

_MAX_ITERATIONS = 200  # 安全上限。ノード数を大幅に超える反復は実装不備の兆候。


@dataclass(frozen=True)
class LevelProposal:
    node_id: str
    current: Optional[int]
    proposed: int
    reason: str


@dataclass(frozen=True)
class ReviewItem:
    node_id: str
    reason: str


@dataclass(frozen=True)
class ProposalResult:
    proposals: List[LevelProposal] = field(default_factory=list)
    needs_review: List[ReviewItem] = field(default_factory=list)
    skipped_multi_dest: List[str] = field(default_factory=list)
    cancelled: bool = False


def _is_multi_dest(node: Dict[str, Any]) -> bool:
    return len(node.get("dests_down") or []) > 1 or len(node.get("dests_right") or []) > 1


def build_reverse_index(nodes: Sequence[Dict[str, Any]]) -> Dict[str, List[Tuple[str, str]]]:
    """dest_id -> [(source_id, direction)] の逆引きを構築する。

    direction は "down"（継承）または "right"（+1）。複数宛先を持つ行は
    ソースとしてもスコープ外のため登録しない（呼び出し側で skipped_multi_dest に計上）。
    """
    reverse: Dict[str, List[Tuple[str, str]]] = {}
    for node in nodes:
        if _is_multi_dest(node):
            continue
        for dest_id in node.get("dests_down") or []:
            reverse.setdefault(dest_id, []).append((node["id"], "down"))
        for dest_id in node.get("dests_right") or []:
            reverse.setdefault(dest_id, []).append((node["id"], "right"))
    return reverse


def _describe(source_id: str, direction: str) -> str:
    if direction == "right":
        return f"行{source_id}から+1(右)"
    return f"行{source_id}を継承(下)"


def _derive_candidate(
    node_id: str,
    reverse_index: Dict[str, List[Tuple[str, str]]],
    values: Dict[str, int],
    conflicted: Dict[str, str],
) -> Tuple[Optional[int], Optional[str], bool]:
    """settled/values を根拠に node_id の推測値を1回だけ導出する（再帰なし）。

    Returns:
        (候補値 or None, 根拠テキスト or None, sources_pending)
        sources_pending=True はソースの一部がまだ未確定であることを示す。
    """
    sources = reverse_index.get(node_id, [])
    if not sources:
        return None, None, False

    pending = any(s not in values and s not in conflicted for s, _ in sources)
    if pending:
        return None, None, True

    candidates: Dict[int, List[str]] = {}
    descriptions: Dict[int, str] = {}
    for source_id, direction in sources:
        if source_id in conflicted:
            continue  # 矛盾したソースからは値を引き継がない
        value = values[source_id] + (1 if direction == "right" else 0)
        candidates.setdefault(value, []).append(source_id)
        descriptions.setdefault(value, _describe(source_id, direction))

    if not candidates:
        return None, None, False  # ソース全滅（全て矛盾）→ 呼び出し側で要確認へ

    if len(candidates) > 1:
        reason = "、".join(
            f"{descriptions[val]}({','.join(srcs)})" for val, srcs in sorted(candidates.items())
        )
        return None, f"矛盾する推測: {reason}", False

    (value,) = candidates.keys()
    return value, descriptions[value], False


def _fixed_point_fill(
    target_ids: Sequence[str],
    reverse_index: Dict[str, List[Tuple[str, str]]],
    known: Dict[str, int],
    stop_event: Any,
) -> Tuple[Dict[str, int], Dict[str, str], Dict[str, str]]:
    """known をアンカーに、target_ids を固定点反復で埋める（再帰禁止）。

    「確定済みの値から推測できるセルだけを1パスで埋める」を、これ以上埋まらなくなる
    まで繰り返す。循環内で空欄同士が連鎖し確定不能なグループは conflicts に
    「推測不能」として残る（doc: 安全性の設計 — 循環対応）。

    Returns:
        settled: known を含む確定値（target_ids のうち導出できたもの）
        conflicts: 矛盾または推測不能だった target_id -> 理由
        reasons: settled の各 target_id -> 根拠テキスト
    """
    settled: Dict[str, int] = dict(known)
    conflicts: Dict[str, str] = {}
    reasons: Dict[str, str] = {}
    unresolved = [nid for nid in target_ids if nid not in settled]

    for _ in range(_MAX_ITERATIONS):
        if stop_event is not None and stop_event.is_set():
            break
        changed = False
        still_unresolved: List[str] = []
        for nid in unresolved:
            value, note, pending = _derive_candidate(nid, reverse_index, settled, conflicts)
            if pending:
                still_unresolved.append(nid)
                continue
            if value is not None:
                settled[nid] = value
                reasons[nid] = note or ""
                changed = True
                continue
            if note is not None:
                conflicts[nid] = note
                changed = True
                continue
            # ソースが無い、または全ソースが矛盾済み → このパスでは確定不能
            still_unresolved.append(nid)
        unresolved = still_unresolved
        if not changed or not unresolved:
            break

    for nid in unresolved:
        conflicts.setdefault(
            nid,
            "推測不能: 参照元の値が確定しないため列を推測できません"
            "（先頭行などはあらかじめ列を入力してください）",
        )

    return settled, conflicts, reasons


def compute_level_proposals(
    nodes: Sequence[Dict[str, Any]],
    raw_levels: Dict[str, Optional[int]],
    *,
    mode: ProposalMode = "blank_only",
    stop_event: Any = None,
) -> ProposalResult:
    """列(level) の提案を計算する（Excel へは一切書き込まない・読み取り専用）。

    Args:
        nodes: parse_table_rows() が返す nodes と同形（id/dests_down/dests_right を使用）。
        raw_levels: node_id -> 列セルの生値（空欄/不正値は None。parse_level_optional 経由で作る）。
        mode: "blank_only"（空欄のみ提案・既定）または "full_recalc"（既存値も対象に含める）。
        stop_event: threading.Event 互換。計算中にセットされたら打ち切り、cancelled=True を返す。
    """
    if stop_event is not None and stop_event.is_set():
        return ProposalResult(cancelled=True)

    multi_dest_ids = [n["id"] for n in nodes if _is_multi_dest(n)]
    eligible_nodes = [n for n in nodes if not _is_multi_dest(n)]
    reverse_index = build_reverse_index(nodes)

    known = {nid: v for nid, v in raw_levels.items() if v is not None}
    all_ids = [n["id"] for n in eligible_nodes]
    blank_ids = [nid for nid in all_ids if nid not in known]

    settled, conflicts, reasons = _fixed_point_fill(blank_ids, reverse_index, known, stop_event)
    if stop_event is not None and stop_event.is_set():
        return ProposalResult(cancelled=True)

    proposals: List[LevelProposal] = []
    needs_review: List[ReviewItem] = []

    for nid in blank_ids:
        if nid in settled:
            proposals.append(
                LevelProposal(
                    node_id=nid, current=None, proposed=settled[nid], reason=reasons[nid]
                )
            )
        elif nid in conflicts:
            needs_review.append(ReviewItem(node_id=nid, reason=conflicts[nid]))

    if mode == "full_recalc":
        for nid in all_ids:
            if nid in blank_ids:
                continue  # 上ですでに処理済み
            if stop_event is not None and stop_event.is_set():
                return ProposalResult(cancelled=True)
            candidate, note, pending = _derive_candidate(nid, reverse_index, settled, conflicts)
            if pending or candidate is None:
                if note is not None:
                    needs_review.append(ReviewItem(node_id=nid, reason=note))
                continue  # ソース無し（エントリ行等）→ 既存値のまま・提案一覧にも出ない
            if candidate != known.get(nid):
                proposals.append(
                    LevelProposal(
                        node_id=nid, current=known.get(nid), proposed=candidate, reason=note or ""
                    )
                )

    return ProposalResult(
        proposals=proposals,
        needs_review=needs_review,
        skipped_multi_dest=multi_dest_ids,
    )
