"""C-2「更新」— 鮮度チェック後に列(level) を Excel へ一括書き込みする。

「提案の計算」（level_inference.compute_level_proposals）は読み取り専用で
Excel を一切変更しない。本モジュールは「更新」ボタンが押されたときだけ呼ばれ、
書き込み直前に Excel を再読取して鮮度チェック（区別①②③）を行ってから、
対象列を Range 一括（2次元配列）で単一 COM 呼び出しにより書き込む。
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Any, Dict, FrozenSet, List, Optional

import pythoncom
import pywintypes

from app.core.excel_engine import get_excel_app
from app.core.level_inference import LevelProposal, ProposalMode
from app.core.live_preview import read_watched_range
from app.core.proposal_fingerprint import (
    LEVEL_COL,
    ProposalSnapshot,
    capture_snapshot,
    hand_edited_ids,
    topology_changed,
)

logger = logging.getLogger("flowchart-excel")


@dataclass(frozen=True)
class WriteResult:
    ok: bool
    updated_count: int = 0
    excluded_count: int = 0
    stale_topology: bool = False
    error: Optional[str] = None


def _norm_id(value: Any) -> str:
    if value is None or value == "":
        return ""
    return str(value).split(".")[0].strip()


def write_level_updates(
    watch: Dict[str, Any],
    proposals: List[LevelProposal],
    *,
    mode: ProposalMode,
    baseline: ProposalSnapshot,
) -> WriteResult:
    """提案されたセルへ、鮮度確認のうえ Excel を一括書き込みする。

    区別①（ID追加削除・接続先(下)/(右)の変更）を検知した場合は書き込みを行わず
    `stale_topology=True` を返す（呼び出し側が再計算・再確認フローへ合流させ、
    自動再試行の上限回数を管理する）。
    区別②（「空欄のみ更新」モード限定の対象セル手入力）は書き込み対象から静かに
    除外するが、件数は `excluded_count` で返し完了メッセージへの明示に使う。
    区別③（Text1〜3・色などの無関係な変更）は判定対象外のため更新をブロックしない。
    """
    if not proposals:
        return WriteResult(ok=True, updated_count=0, excluded_count=0)

    pythoncom.CoInitialize()
    try:
        data, _ = read_watched_range(watch)
        fresh = capture_snapshot(data)

        if topology_changed(baseline, fresh):
            return WriteResult(ok=False, stale_topology=True)

        excluded: FrozenSet[str] = (
            hand_edited_ids(baseline, fresh) if mode == "blank_only" else frozenset()
        )
        to_write = {p.node_id: p.proposed for p in proposals if p.node_id not in excluded}
        excluded_count = sum(1 for p in proposals if p.node_id in excluded)

        if not to_write:
            return WriteResult(ok=True, updated_count=0, excluded_count=excluded_count)

        app = get_excel_app()
        if not app:
            return WriteResult(ok=False, error="Excelが起動していません。")

        workbook = None
        for wb in app.Workbooks:
            if str(wb.Name) == str(watch.get("workbookName")):
                workbook = wb
                break
        if workbook is None:
            return WriteResult(ok=False, error="ブックが見つかりません。")

        sheet = workbook.Sheets(watch.get("sheetName"))
        is_full = bool(watch.get("isFullMode"))
        if is_full:
            r_tgt = sheet.Range(watch["anchorAddress"]).CurrentRegion
        else:
            addr = watch.get("rangeAddress") or watch["anchorAddress"]
            r_tgt = sheet.Range(addr)

        row_count = len(data)

        # ID→行位置を直前の再読取結果から解決し直す（COM の Range 書き込みは
        # 行位置ベースのため、書き込み直前に安全確認する TOCTOU 窓を塞ぐ保険）。
        id_to_row: Dict[str, int] = {}
        for idx, row in enumerate(data):
            nid = _norm_id(row[0] if row else None)
            if nid:
                id_to_row[nid] = idx

        if any(nid not in id_to_row for nid in to_write):
            # ID再解決に失敗（行の移動・削除等）→ 区別①と同じ自動再計算フローに合流
            return WriteResult(ok=False, stale_topology=True)

        level_range = r_tgt.Cells(1, LEVEL_COL + 1).Resize(row_count, 1)
        current_values = level_range.Value
        column: List[List[Any]] = (
            [[current_values]] if row_count == 1 else [[row[0]] for row in current_values]
        )
        for nid, new_value in to_write.items():
            column[id_to_row[nid]][0] = new_value

        level_range.Value = column
    except (pywintypes.com_error, AttributeError, RuntimeError) as exc:
        logger.exception("level_update_write_failed")
        return WriteResult(ok=False, error=f"更新に失敗しました: {exc}")
    finally:
        pythoncom.CoUninitialize()

    return WriteResult(ok=True, updated_count=len(to_write), excluded_count=excluded_count)
