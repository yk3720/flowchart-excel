"""F5「更新」— 鮮度チェック後にID列を Excel へ書き込む（構想設計§6-2）。

`level_writer.write_level_updates`と同じ「統一読み取り→書き込み」のTOCTOU最小化
パターン・`capture_snapshot`/`topology_changed`による既存行の鮮度チェックを再利用する。

一方、対象行の再特定方法は`write_level_updates`とは異なる。`write_level_updates`は
既存のID値で書き込み先の行を再特定できるが、F5が対象にする行は**ID自体がまだ空欄**
のため、同じ仕組みでは再特定できない（直接の関数再利用が構造的に不可能という、構想
設計レビュー時には見落とされていた前提）。代わりに、提案時点の行位置(`row_index`)と
「その行が今もID空欄のままであること」、および監視対象テーブル全体の行数が提案計算
時点から変化していないことを、書き込み直前に再確認する。いずれかが崩れていれば
`stale_topology=True`を返し、既存のC-2と同じ自動再計算・再確認フローに合流させる。

書き込み自体も`write_level_updates`の「列全体への一括 Value 代入」とは異なる。雛形は
タイトル行が ListObject の直上に接するため CurrentRegion がタイトル＋表に広がり、
ID列全体への一括代入が Excel に反映されない（例外なし）実機不具合があった。そのため
提案行だけセル単位で書き、直後の読み戻しで反映を検証する。
"""
from __future__ import annotations

import logging

import pythoncom
import pywintypes

from app.core.excel_engine import get_excel_app
from app.core.id_proposal import ID_COL, IdProposal
from app.core.level_writer import WriteResult
from app.core.live_preview import read_watched_range
from app.core.parse_table import norm_id
from app.core.proposal_fingerprint import (
    ProposalSnapshot,
    capture_snapshot,
    topology_changed,
)

logger = logging.getLogger("flowchart-excel")


def write_id_updates(
    watch: dict,
    proposals: tuple[IdProposal, ...],
    *,
    baseline_row_count: int,
    baseline: ProposalSnapshot,
) -> WriteResult:
    """提案されたID値を、鮮度確認のうえ Excel へ一括書き込みする。

    既存行（IDを持つ行）の鮮度は`topology_changed`で、新規行を含むテーブル全体の
    行数変化は`baseline_row_count`との比較で検知する（行数が1でも変わっていれば
    安全側で中断する。複数の新規行が同時に出現・消失するケースを含めて位置の
    ズレを前提にしないため）。
    """
    if not proposals:
        return WriteResult(ok=True, updated_count=0, excluded_count=0)

    pythoncom.CoInitialize()
    try:
        data, _ = read_watched_range(watch)
        fresh = capture_snapshot(data)

        if topology_changed(baseline, fresh) or len(data) != baseline_row_count:
            return WriteResult(ok=False, stale_topology=True)

        app = get_excel_app()
        if not app:
            return WriteResult(ok=False, error="Excelが起動していません。")

        app.ScreenUpdating = False
        app.DisplayAlerts = False
        try:
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

            # 統一読み取り（書き込み直前の2回目の読み取りを行わない。level_writerと同じ方針）。
            fresh_data = r_tgt.Value
            if not fresh_data or not isinstance(fresh_data, tuple):
                return WriteResult(ok=False, stale_topology=True)

            row_count = len(fresh_data)
            if row_count != baseline_row_count or int(r_tgt.Rows.Count) != row_count:
                return WriteResult(ok=False, stale_topology=True)

            # 対象行が今もID空欄のままであることを再確認する（1行でも崩れていれば中断）。
            for p in proposals:
                if p.row_index >= row_count:
                    return WriteResult(ok=False, stale_topology=True)
                row = fresh_data[p.row_index]
                if norm_id(row[ID_COL] if row else None):
                    return WriteResult(ok=False, stale_topology=True)

            # 雛形はタイトル行が表（ListObject）の直上に接するため CurrentRegion が
            # タイトル＋ヘッダー＋データに広がる。ID列全体への一括 Value 代入は、
            # この混在 Range では Excel 側に反映されない（例外も出ない）ことが
            # 実機で確認された。提案行だけセル単位で書き、直後に読み戻して検証する。
            for p in proposals:
                r_tgt.Cells(p.row_index + 1, ID_COL + 1).Value = p.proposed_id

            verify = r_tgt.Value
            if not verify or not isinstance(verify, tuple) or len(verify) != row_count:
                return WriteResult(
                    ok=False,
                    error="ID採番の書き込み後に表を再読取できませんでした。もう一度お試しください。",
                )
            for p in proposals:
                row = verify[p.row_index]
                got = norm_id(row[ID_COL] if row else None)
                if got != norm_id(p.proposed_id):
                    logger.error(
                        "id_update_verify_failed | row_index=%s | expected=%s | got=%s",
                        p.row_index,
                        p.proposed_id,
                        got,
                    )
                    return WriteResult(
                        ok=False,
                        error=(
                            "ID列への書き込みがExcelに反映されませんでした。"
                            "表のセルを選択し直してから、もう一度「提案の計算」→「更新」を試してください。"
                        ),
                    )
        finally:
            app.ScreenUpdating = True
            app.DisplayAlerts = True
    except (pywintypes.com_error, AttributeError, RuntimeError) as exc:
        logger.exception("id_update_write_failed")
        return WriteResult(ok=False, error=f"ID採番の書き込みに失敗しました: {exc}")
    finally:
        pythoncom.CoUninitialize()

    return WriteResult(ok=True, updated_count=len(proposals), excluded_count=0)
