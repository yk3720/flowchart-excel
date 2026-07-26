"""flowchart-excel 定数定義。

10列（table-10col-v2）— SSOT: flowchart-studio/lib/flowchart/table/tableColumns.ts
"""
import platform
from typing import Any, Dict, List

# --- アプリケーション情報 ---
APP_NAME = "Flowchart Excel"
REVISION = "rev005"
AUTHOR = "YK"
RECOMMENDED_PYTHON = "3.14.2"

# --- 10列 v2 ヘッダー ---
TABLE_HEADERS_10_V2: List[str] = [
    "ID",
    "図形種別",
    "色",
    "接続先(下)",
    "接続先(右)",
    "段",
    "列",
    "Text1",
    "Text2",
    "Text3",
]


class ExcelConstants:
    """Excel COM 定数。"""

    MSOSHAPE_RECTANGLE = 1
    MSOSHAPE_DIAMOND = 4
    MSOSHAPE_PARALLELOGRAM = 2
    MSOSHAPE_ROUNDED_RECTANGLE = 5
    MSOSHAPE_MANUAL_INPUT = 11

    MSOCONNECTOR_STRAIGHT = 1
    MSOCONNECTOR_ELBOW = 2

    MSO_ALIGN_CENTER = 2
    MSO_ANCHOR_MIDDLE = 3
    XL_CENTER = -4108
    XL_RIGHT = -4152

    CONNECTOR_SITE_TOP = 1
    CONNECTOR_SITE_LEFT = 2
    CONNECTOR_SITE_BOTTOM = 3
    CONNECTOR_SITE_RIGHT = 4

    TITLE_BG_COLOR = 0xFFE4B5
    HEADER_BG_COLOR = 0xD7D7D7


FONT_FAMILY = "Yu Gothic UI" if platform.system() == "Windows" else "Arial"
APP_FONT = (FONT_FAMILY, 11)
TITLE_FONT = (FONT_FAMILY, 14, "bold")
LABEL_FONT = (FONT_FAMILY, 12, "bold")
SMALL_FONT = (FONT_FAMILY, 10)

# --- デザイントークン（rev015: flowchart-studio 準拠） ---
# SSOT: flowchart-studio/app/globals.css の --flow-* 変数（Layer C 操作 chrome）。
# 値を変える場合は両リポで同期する。
FLOW_ACCENT = "#2563eb"
FLOW_ACCENT_HOVER = "#1d4ed8"
FLOW_SURFACE = "#ffffff"
FLOW_SURFACE_MUTED = "#f8fafc"
FLOW_SURFACE_SUBTLE = "#f1f5f9"
FLOW_BORDER = "#e2e8f0"
FLOW_BORDER_STRONG = "#cbd5e1"
FLOW_TEXT = "#0f172a"
FLOW_TEXT_BODY = "#334155"
FLOW_TEXT_MUTED = "#64748b"
FLOW_DANGER = "#dc2626"
FLOW_DANGER_HOVER = "#b91c1c"
FLOW_SUCCESS_BG = "#f0fdf4"
FLOW_SUCCESS_BORDER = "#bbf7d0"
FLOW_SUCCESS_TEXT = "#14532d"

# 枠線太さ・角丸は画面内で単一値に統一する（VISUAL_DESIGN_RULES §2 SSOT）
CARD_BORDER_WIDTH = 1
CORNER_RADIUS = 8

THEMES: Dict[str, Dict[str, int]] = {
    "標準（信頼）": {"connector": 0xC07000, "shape_line": 0x000000, "label": 0x000000},
    "安全（正常）": {"connector": 0x50B000, "shape_line": 0x306000, "label": 0x000000},
    "警告（注意）": {"connector": 0x00C0FF, "shape_line": 0x0000FF, "label": 0x000000},
}

DEFAULT_BOX_HEIGHT = 60.0
DEFAULT_BOX_WIDTH = 160.0
DEFAULT_GAP_V = 30.0
DEFAULT_GAP_H = 100.0

PRESETS = [
    {"id": "large", "name": "大", "height": 100.0, "width": 220.0, "gap_v": 40.0, "gap_h": 150.0},
    {"id": "medium", "name": "中", "height": 60.0, "width": 160.0, "gap_v": 30.0, "gap_h": 100.0},
    {"id": "small", "name": "小", "height": 45.0, "width": 120.0, "gap_v": 20.0, "gap_h": 80.0},
]


def _legacy_8_to_10_v2(rows_8: List[List[Any]]) -> List[List[Any]]:
    """旧 8 列雛形を table-10col-v2 に変換（段=行番号 · 列=旧 Level）。"""
    converted: List[List[Any]] = []
    for tier, row in enumerate(rows_8):
        converted.append(
            [
                row[0],
                row[1],
                "",
                row[2],
                row[3],
                tier,
                row[4],
                row[5],
                row[6],
                row[7] if len(row) > 7 else "",
            ]
        )
    return converted


_LEGACY_TEMPLATE_8: Dict[str, List[List[Any]]] = {
    "simple_no": [
        [10, "端子", "20", "", 0, "開始", "カレー作り開始", ""],
        [20, "処理", "30", "", 0, "手を洗う", "衛生管理の徹底", ""],
        [30, "処理", "40", "", 0, "材料を揃える", "肉、玉ねぎ、人参、ポテト", ""],
        [40, "処理", "50", "", 0, "野菜を切る", "一口サイズに", ""],
        [50, "処理", "60", "", 0, "肉を炒める", "色が変わるまで", ""],
        [60, "処理", "70", "", 0, "野菜を加えて炒める", "油が回るまで", ""],
        [70, "処理", "80", "", 0, "水を入れ煮込む", "中火で20分", ""],
        [80, "処理", "90", "", 0, "アクを取る", "丁寧に", ""],
        [90, "処理", "100", "", 0, "ルーを入れる", "一旦火を止める", ""],
        [100, "処理", "110", "", 0, "弱火で煮込む", "とろみがつくまで", ""],
        [110, "処理", "120", "", 0, "盛り付け", "ご飯とカレー", ""],
        [120, "端子", "", "", 0, "終了", "美味しいカレーの完成", ""],
    ],
    "simple_yes": [
        [10, "端子", "20", "", 0, "開始", "", ""],
        [20, "処理", "30", "", 0, "下準備（野菜・肉）", "", ""],
        [30, "処理", "40", "", 0, "炒める・煮込む", "", ""],
        [40, "判断", "60", "50", 0, "野菜は柔らかい？", "", ""],
        [50, "処理", "30", "", 1, "数分追加で煮込む", "再試行", ""],
        [60, "処理", "70", "", 0, "カレールーを入れる", "", ""],
        [70, "判断", "90", "80", 0, "味は丁度良い？", "", ""],
        [80, "処理", "70", "", 1, "調味料を追加", "ソース、ケチャップ等", ""],
        [90, "処理", "100", "", 0, "ひと煮立ちさせる", "", ""],
        [100, "判断", "120", "110", 0, "辛さは大丈夫？", "", ""],
        [110, "処理", "90", "", 1, "牛乳や蜂蜜を足す", "辛さ緩和", ""],
        [120, "処理", "130", "", 0, "お皿に盛り付ける", "", ""],
        [130, "端子", "", "", 0, "終了", "", ""],
    ],
    "complex_no": [
        [10, "端子", "20", "", 0, "調理開始", "", ""],
        [20, "処理", "30", "", 0, "お米を研ぐ", "", ""],
        [30, "処理", "40", "", 0, "炊飯スイッチON", "45分", ""],
        [40, "処理", "50", "", 1, "野菜を切る", "カレー用", ""],
        [50, "処理", "60", "", 1, "肉を炒める", "", ""],
        [60, "処理", "70", "", 1, "野菜を入れ炒める", "", ""],
        [70, "処理", "80", "", 1, "煮込む・アク取り", "", ""],
        [80, "処理", "90", "", 1, "ルーを溶かす", "", ""],
        [90, "処理", "100", "", 0, "サラダを作る", "サイドメニュー", ""],
        [100, "処理", "110", "", 0, "ドレッシング準備", "", ""],
        [110, "処理", "120", "", 0, "炊き上がり確認", "蒸らし完了", ""],
        [120, "処理", "130", "", 0, "全部盛り付け", "", ""],
        [130, "処理", "140", "", 0, "後片付け", "食器洗い", ""],
        [140, "端子", "", "", 0, "終了", "", ""],
    ],
    "complex_yes": [
        [10, "端子", "20", "", 0, "究極カレー開始", "", ""],
        [20, "処理", "30", "", 0, "玉ねぎを飴色にする", "約30分", ""],
        [30, "処理", "40", "", 0, "スパイスを調合", "独自ブレンド", ""],
        [40, "処理", "50", "", 0, "肉を焼き固める", "旨味を閉じ込める", ""],
        [50, "処理", "60", "", 0, "赤ワインで煮込む", "", ""],
        [60, "判断", "80", "70", 0, "水分量は適正？", "", ""],
        [70, "処理", "50", "", 1, "ブイヨンを足す", "調整", ""],
        [80, "処理", "90", "", 0, "ルーと隠し味投入", "チョコ、コーヒー", ""],
        [90, "判断", "110", "100", 0, "深みは出ている？", "", ""],
        [100, "処理", "80", "", 1, "醤油を一垂らし", "調整", ""],
        [110, "判断", "130", "120", 0, "一晩寝かせる？", "", ""],
        [120, "処理", "110", "", 1, "粗熱を取る", "冷蔵庫へ", ""],
        [130, "処理", "140", "", 0, "温め直して盛り付け", "", ""],
        [140, "処理", "150", "", 0, "完成披露", "実食", ""],
        [150, "端子", "", "", 0, "終了", "", ""],
    ],
}

TEMPLATE_DATA: Dict[str, List[List[Any]]] = {
    key: _legacy_8_to_10_v2(rows) for key, rows in _LEGACY_TEMPLATE_8.items()
}
