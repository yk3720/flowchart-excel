"""C-2 更新直前の鮮度チェック用スナップショット（表全体ではなく提案の根拠だけを追跡）。

`live_preview.table_fingerprint()` は表全体（title/schema/table/layout）を丸ごと
JSON シリアライズするため、Text1〜3 等の無関係な列の変更でも「表が変わった」扱いに
なってしまい C-2 には粗すぎる（区別③の変更まで区別①として弾いてしまう）。
本モジュールは ID・接続先(下)/(右)・提案対象セルだけを対象にした軽量な比較を提供する。

table-10col-v2 の列位置（parse_table.TABLE_HEADERS_10_V2 と同順）を既定値に持つが、
列インデックスは引数化してあり、将来 C-3（段の再採番）が target_col=TIER_COL で
そのまま再利用できる。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, FrozenSet, Tuple

from app.core.parse_table import norm_id

ID_COL = 0
DEST_DOWN_COL = 3
DEST_RIGHT_COL = 4
LEVEL_COL = 6  # C-2 の対象列。C-3 では TIER_COL(=5) を target_col に渡す想定。
TIER_COL = 5


@dataclass(frozen=True)
class ProposalSnapshot:
    ids: FrozenSet[str]
    edges: Dict[str, Tuple[str, str]]  # id -> (raw 接続先(下)テキスト, raw 接続先(右)テキスト)
    target_cells: Dict[str, Any]  # id -> raw 提案対象セル値（C-2 では列セル）


def capture_snapshot(
    rows_raw: Any,
    *,
    id_col: int = ID_COL,
    down_col: int = DEST_DOWN_COL,
    right_col: int = DEST_RIGHT_COL,
    target_col: int = LEVEL_COL,
) -> ProposalSnapshot:
    """生の Excel 行データ（タプルのタプル）から鮮度チェック用スナップショットを作る。"""
    ids: set[str] = set()
    edges: Dict[str, Tuple[str, str]] = {}
    target_cells: Dict[str, Any] = {}
    for row in rows_raw:
        nid = norm_id(row[id_col] if len(row) > id_col else None)
        if not nid:
            continue
        ids.add(nid)
        down_raw = row[down_col] if len(row) > down_col else None
        right_raw = row[right_col] if len(row) > right_col else None
        edges[nid] = (
            "" if down_raw is None else str(down_raw),
            "" if right_raw is None else str(right_raw),
        )
        target_cells[nid] = row[target_col] if len(row) > target_col else None
    return ProposalSnapshot(ids=frozenset(ids), edges=edges, target_cells=target_cells)


def topology_changed(old: ProposalSnapshot, new: ProposalSnapshot) -> bool:
    """ID の追加・削除、または接続先(下)/(右)の生テキストが変わっていれば True（区別①）。"""
    if old.ids != new.ids:
        return True
    return any(old.edges.get(nid) != new.edges.get(nid) for nid in old.ids)


def hand_edited_ids(old: ProposalSnapshot, new: ProposalSnapshot) -> FrozenSet[str]:
    """提案対象セル自体が手入力で変わった行の ID 集合（区別②）。

    「空欄のみ更新」モード限定で呼び出し側が使う判定材料。全部再計算モードでは
    区別②を適用しない（呼び出し側の責務。本関数はモードに関知しない）。
    """
    changed = {
        nid for nid in old.ids & new.ids if old.target_cells.get(nid) != new.target_cells.get(nid)
    }
    return frozenset(changed)
