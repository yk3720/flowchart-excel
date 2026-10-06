"""F7: Excel作成・更新の実行前にAutoSave状態を確認し、無効なら軽いリマインダーを返す。

完全なロールバック（Excelプロセス死亡後にアプリ側で状態を復元する）は、プロセスが
死んでいればCOM操作そのものが成立しないため技術的に不可能。「どこまでの予防導線を
用意すれば十分か」はユーザーのリスク許容度に依存するポリシー選択であり、現時点の
暫定方針（構想設計§7・正式承認不要）として、Excel自身のAutoSave（自動保存/バージョン
履歴）が有効かどうかを確認し、無効なら軽いリマインダーを表示する程度に留める。

却下した代替案: 生成前に独自のスナップショット（ファイルコピー）を取る。却下理由:
ユーザーが別のアプリでExcelファイルを開いている場合に書き込み競合のリスクがあり、
かつExcel自体のAutoSave機能と役割が重複する。
"""
from __future__ import annotations

import logging
from typing import Any

import pywintypes

from app.core.excel_engine import get_excel_app

logger = logging.getLogger("flowchart-excel")

AUTOSAVE_REMINDER = (
    "このブックはAutoSave（自動保存）が無効です。"
    "Excel側で手動保存してから続行することをおすすめします。"
)


def autosave_reminder(workbook: Any) -> str | None:
    """渡されたワークブックのAutoSaveが無効ならリマインダー文言を返す（有効/判定不能ならNone）。"""
    try:
        if bool(workbook.AutoSaveOn):
            return None
    except (pywintypes.com_error, AttributeError):
        # AutoSaveOnプロパティ自体を取得できない（古いExcel・ローカルファイル等）
        # → 判定不能として何も言わない（誤ったリマインダーを出さない）。
        return None
    return AUTOSAVE_REMINDER


def autosave_reminder_for_watch(watch: dict[str, Any]) -> str | None:
    """watchメタからワークブックを解決し、AutoSave状態に応じたリマインダーを返す。"""
    app = get_excel_app()
    if not app:
        return None
    workbook = None
    for wb in app.Workbooks:
        if str(wb.Name) == str(watch.get("workbookName")):
            workbook = wb
            break
    if workbook is None:
        return None
    return autosave_reminder(workbook)
