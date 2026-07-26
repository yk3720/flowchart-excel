"""プレビュー表示中の Excel 再読込（ライブ更新）。"""
from __future__ import annotations

import logging
from typing import Any, Dict, Optional, Tuple

import pythoncom
import pywintypes

from app.constants import ExcelConstants
from app.core.excel_engine import get_excel_app
from app.core.preview_payload import build_studio_preview_payload

logger = logging.getLogger("flowchart-excel")

# ポーリング間隔（秒）— 編集確定後に追従する想定
LIVE_POLL_INTERVAL_SEC = 0.75


def layout_to_config(layout: Dict[str, Any]) -> Dict[str, float]:
    return {
        "width": float(layout["width"]),
        "height": float(layout["heightMin"]),
        "gap_v": float(layout["gapV"]),
        "gap_h": float(layout["gapH"]),
    }


def table_fingerprint(payload: Dict[str, Any]) -> str:
    """表内容の差分検出用（レイアウト変更は別経路）。"""
    import json

    return json.dumps(
        {
            "title": payload.get("title"),
            "schema": payload.get("schema"),
            "table": payload.get("table"),
            "layout": payload.get("layout"),
        },
        ensure_ascii=False,
        sort_keys=True,
        default=str,
    )


def _read_watched_range(watch: Dict[str, Any]) -> Tuple[Any, str]:
    """watch メタから Excel 範囲を再取得する。"""
    app = get_excel_app()
    if not app:
        raise RuntimeError("Excelが起動していません。")

    workbook_name = watch.get("workbookName")
    sheet_name = watch.get("sheetName")
    if not workbook_name or not sheet_name:
        raise ValueError("watch メタが不完全です")

    workbook = None
    for wb in app.Workbooks:
        if str(wb.Name) == str(workbook_name):
            workbook = wb
            break
    if workbook is None:
        raise RuntimeError(f"ブックが見つかりません: {workbook_name}")

    sheet = workbook.Sheets(sheet_name)
    is_full = bool(watch.get("isFullMode"))
    if is_full:
        anchor = sheet.Range(watch["anchorAddress"])
        r_tgt = anchor.CurrentRegion
    else:
        addr = watch.get("rangeAddress") or watch["anchorAddress"]
        r_tgt = sheet.Range(addr)

    data = r_tgt.Value
    if not data or not isinstance(data, tuple):
        raise ValueError("選択範囲にデータがありません。")

    title_txt = "フローチャート"
    start_cell = r_tgt.Cells(1, 1)
    if is_full:
        for i in range(-5, 1):
            row_idx = max(1, int(start_cell.Row) + i)
            cell = sheet.Cells(row_idx, start_cell.Column)
            if cell.Interior.Color == ExcelConstants.TITLE_BG_COLOR and cell.Value:
                title_txt = str(cell.Value)
                break

    return data, title_txt


def try_refresh_studio_payload(base: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Excel を再読込して新ペイロードを返す。失敗・編集中は None。"""
    watch = (base.get("meta") or {}).get("watch")
    if not watch:
        return None

    pythoncom.CoInitialize()
    try:
        data, title_txt = _read_watched_range(watch)
        layout = base.get("layout") or {}
        config = layout_to_config(layout)
        fresh = build_studio_preview_payload(
            data,
            title=title_txt,
            is_full_mode=bool(watch.get("isFullMode")),
            config=config,
        )
        # watch / live フラグを維持
        meta = dict(fresh.get("meta") or {})
        meta["watch"] = watch
        meta["live"] = True
        fresh["meta"] = meta
        return fresh
    except ValueError as exc:
        # ノード0件などはスキップ（直前の表示を維持）
        logger.debug("live_refresh_skip | %s", exc)
        return None
    except (pywintypes.com_error, AttributeError, RuntimeError) as exc:
        # セル編集中など
        logger.debug("live_refresh_com_skip | %s", exc)
        return None
    finally:
        pythoncom.CoUninitialize()
