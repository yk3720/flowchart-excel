"""F4: 新規行へのデータ入力規則（種別・色列）の継承（構想設計§5）。

雛形はExcelネイティブテーブル（`ListObject`）。F2のデータ入力規則は雛形作成時の
固定行にのみ適用され、行挿入・貼り付けで追加された新規行への自動継承は無い。
本モジュールは、ライブプレビューの既存ポーリング（`embedded_preview._live_tick`、
0.75秒間隔）に相乗りし、監視対象テーブルの現在の行数を前回値と比較して、増えた分
（末尾に追加されたと仮定）にだけ入力規則を適用する。

既知の限界: 行数の単純比較のため、同一ポーリング間隔内に削除と追加が同時発生し
行数が相殺される操作（範囲選択での上書き貼り付け等）は検知できない。
"""
from __future__ import annotations

import logging
from typing import Any

import pywintypes

from app.core.excel_engine import get_excel_app

logger = logging.getLogger("flowchart-excel")

# table-10col-v2（parse_table.TABLE_HEADERS_10_V2と同順）の列インデックス（0-based）。
TYPE_COL = 1
COLOR_COL = 2


def apply_new_row_validation(
    watch: dict[str, Any],
    previous_row_count: int | None,
    *,
    type_validation_list: str,
    color_validation_list: str,
) -> int | None:
    """監視対象テーブルの行数が増えていれば、増分の行に入力規則を適用する。

    `watch["isFullMode"]`が`False`（固定アドレスのレンジ監視）の場合は対象外とし
    `previous_row_count`をそのまま返す（固定レンジは行追加を検知できないため）。
    戻り値は次回呼び出しに渡す「現在の行数」（呼び出し側が保持する）。
    `previous_row_count`が`None`（初回呼び出し）の場合は適用せず、基準値として
    現在の行数を返すだけに留める（雛形作成時にF2が既に適用済みの行への重複適用を
    避けるため）。
    """
    if not bool(watch.get("isFullMode")):
        return previous_row_count

    app = get_excel_app()
    if not app:
        return previous_row_count

    workbook = None
    for wb in app.Workbooks:
        if str(wb.Name) == str(watch.get("workbookName")):
            workbook = wb
            break
    if workbook is None:
        return previous_row_count

    try:
        sheet = workbook.Sheets(watch.get("sheetName"))
        r_tgt = sheet.Range(watch["anchorAddress"]).CurrentRegion
        current_row_count = int(r_tgt.Rows.Count)
    except (pywintypes.com_error, AttributeError):
        return previous_row_count

    if previous_row_count is None or current_row_count <= previous_row_count:
        return current_row_count

    added = current_row_count - previous_row_count
    try:
        new_rows = r_tgt.Rows(previous_row_count + 1).Resize(added)

        type_range = new_rows.Cells(1, TYPE_COL + 1).Resize(added, 1)
        type_range.Validation.Delete()
        type_range.Validation.Add(3, 1, 1, type_validation_list)

        color_range = new_rows.Cells(1, COLOR_COL + 1).Resize(added, 1)
        color_range.Validation.Delete()
        color_range.Validation.Add(3, 1, 1, color_validation_list)

        logger.info(
            "f4_new_row_validation_applied | added=%s | total=%s", added, current_row_count
        )
    except (pywintypes.com_error, AttributeError):
        logger.warning("f4_new_row_validation_failed", exc_info=True)

    return current_row_count
