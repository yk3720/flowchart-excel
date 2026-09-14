"""表「色」列 → 塗り。

SSOT: flowchart-studio/lib/flowchart/visual/flowColors.ts
Excel COM の Fill.ForeColor.RGB は VBA RGB（R + G*256 + B*65536）。
"""
from __future__ import annotations

FILL_HEX_BY_KEYWORD: dict[str, str] = {
    "黄": "#fef9c3",
    "橙": "#ffedd5",
    "青": "#dbeafe",
}
DEFAULT_FILL_HEX = "#ffffff"


def hex_to_vba_rgb(hex_color: str) -> int:
    """#RRGGBB を Excel COM 用の VBA RGB 整数にする。"""
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return hex_to_vba_rgb(DEFAULT_FILL_HEX)
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return int(r + (g << 8) + (b << 16))


def fill_hex_for_hint(raw: str | None) -> str:
    """表の色セル → プレビュー用 #RRGGBB。未知値（赤など）は通常の白。"""
    if raw is None:
        return DEFAULT_FILL_HEX
    key = str(raw).strip()
    if not key:
        return DEFAULT_FILL_HEX
    return FILL_HEX_BY_KEYWORD.get(key, DEFAULT_FILL_HEX)


def fill_vba_rgb_for_hint(raw: str | None) -> int:
    """表の色セル → Excel AutoShape の Fill.ForeColor.RGB。"""
    return hex_to_vba_rgb(fill_hex_for_hint(raw))
