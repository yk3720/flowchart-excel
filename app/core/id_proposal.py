"""F5: 新規行（ID空欄）へのID自動採番の提案（構想設計§6、読み取り専用）。

接続先（配線）の自動追随は行わない（構想設計§6-3でスコープ外・ユーザー承認済み）。
ここで提案するのは新規行へのIDの値のみで、Excelへの書き込みは行わない
（書き込みは `app.core.id_writer.write_id_updates` が「更新」ボタン操作時にのみ行う）。

採番基準は「シート上に現存する最大ID+1」（要承認事項#2・方向性B）。削除済みIDの
再利用自体は許容し、採番後に既存のダングリング参照（現存しない行IDへの接続先(下)/(右)
の生テキスト参照）と値が一致する場合だけ警告する（履歴の永続化は行わない軽量な方式。
方向性A「削除済みIDも含めた履歴管理」は不採用）。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from app.core.parse_table import norm_id

ID_COL = 0
DEST_DOWN_COL = 3
DEST_RIGHT_COL = 4


@dataclass(frozen=True)
class IdProposal:
    row_index: int  # watch範囲（生データ配列）内での0-based行位置。書き込み時の対象行再特定に使う。
    proposed_id: str
    reason: str


@dataclass(frozen=True)
class IdProposalResult:
    proposals: tuple[IdProposal, ...]
    # 採番値が既存のダングリング参照（削除済み行への生テキスト参照）と一致した場合の警告。
    dangling_warnings: tuple[str, ...]


def _is_blank_row(row: Any) -> bool:
    return not any(
        row[c] not in (None, "") for c in range(len(row)) if c != ID_COL
    )


def compute_id_proposals(rows: Any) -> IdProposalResult:
    """ID列が空欄の行に「現存する最大ID+1」からの連番を提案する。

    複数の新規行が同一バッチに存在する場合、Excelの行出現順（上から下）に重複しない
    連番を割り当てる（構想設計§6-2「守るべき方針」）。ID以外の列も全て空欄の行
    （実体の無いプレースホルダ行）は対象外とする。
    """
    existing_ids: set[str] = set()
    existing_numeric_ids: set[int] = set()
    for row in rows:
        nid = norm_id(row[ID_COL] if row else None)
        if not nid:
            continue
        existing_ids.add(nid)
        if nid.isdigit():
            existing_numeric_ids.add(int(nid))

    next_id = (max(existing_numeric_ids) if existing_numeric_ids else 0) + 1

    proposals: list[IdProposal] = []
    for idx, row in enumerate(rows):
        nid = norm_id(row[ID_COL] if row else None)
        if nid or _is_blank_row(row):
            continue
        proposals.append(
            IdProposal(
                row_index=idx,
                proposed_id=str(next_id),
                reason="新規行（既存の最大ID+1）",
            )
        )
        next_id += 1

    if not proposals:
        return IdProposalResult(proposals=(), dangling_warnings=())

    dangling: set[str] = set()
    for row in rows:
        for col in (DEST_DOWN_COL, DEST_RIGHT_COL):
            ref = norm_id(row[col] if len(row) > col else None)
            if ref and ref not in existing_ids:
                dangling.add(ref)

    assigned_ids = {p.proposed_id for p in proposals}
    hit = sorted(assigned_ids & dangling, key=int)
    warnings = tuple(
        f"ID {aid} は削除済みの行への参照が残っていたため、新規行への採番により"
        f"意図せず接続される可能性があります。接続先(下)/(右)列を確認してください。"
        for aid in hit
    )
    return IdProposalResult(proposals=tuple(proposals), dangling_warnings=warnings)
