"""フロープレビュー確認ダイアログ（作成前レビュー）。"""
from __future__ import annotations

import tkinter as tk
from typing import Callable

import customtkinter as ctk

from app.constants import (
    APP_FONT,
    CARD_BORDER_WIDTH,
    CORNER_RADIUS,
    FLOW_ACCENT,
    FLOW_ACCENT_HOVER,
    FLOW_BORDER,
    FLOW_DANGER,
    FLOW_SURFACE,
    FLOW_SURFACE_MUTED,
    FLOW_TEXT,
    FLOW_TEXT_BODY,
    FLOW_TEXT_MUTED,
    FONT_FAMILY,
    LABEL_FONT,
    SMALL_FONT,
)
from app.core.layout_preview import PlacedNode, PreviewModel

# 枠線太さ SSOT（同一プレビュー内で統一 · VISUAL_DESIGN_RULES）
PREVIEW_STROKE = 2
PREVIEW_PAD = 24


def _rgb_hex(bgr: int) -> str:
    """Excel BGR 整数を #RRGGBB に変換する。"""
    r = bgr & 0xFF
    g = (bgr >> 8) & 0xFF
    b = (bgr >> 16) & 0xFF
    return f"#{r:02x}{g:02x}{b:02x}"


class FlowPreviewDialog(ctk.CTkToplevel):
    """プレビュー必須: 確認後にのみ Excel 作成を許可する。"""

    def __init__(
        self,
        master: tk.Misc,
        model: PreviewModel,
        *,
        shape_line_bgr: int,
        connector_bgr: int,
        on_confirm: Callable[[], None],
    ) -> None:
        super().__init__(master)
        self.title("フロープレビュー — 確認してから作成")
        self.configure(fg_color=FLOW_SURFACE_MUTED)
        self.geometry("920x640")
        self.minsize(640, 480)
        self.transient(master)
        self.grab_set()
        self.attributes("-topmost", True)

        self._model = model
        self._on_confirm = on_confirm
        self._shape_line = _rgb_hex(shape_line_bgr)
        self._connector = _rgb_hex(connector_bgr)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_canvas()
        self._build_footer()
        self.after(50, self._draw)
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _build_header(self) -> None:
        head = ctk.CTkFrame(
            self,
            fg_color=FLOW_SURFACE,
            border_width=CARD_BORDER_WIDTH,
            border_color=FLOW_BORDER,
            corner_radius=CORNER_RADIUS,
        )
        head.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))
        head.grid_columnconfigure(0, weight=1)

        mode = "表全体" if self._model.is_full_mode else "選択範囲"
        ctk.CTkLabel(
            head,
            text=f"{self._model.title}（{mode}）",
            font=LABEL_FONT,
            text_color=FLOW_TEXT,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 2))

        ctk.CTkLabel(
            head,
            text=(
                f"ノード {self._model.node_count} · 接続 {self._model.edge_count}"
                "  —  問題なければ「Excelに作成」を押してください"
            ),
            font=SMALL_FONT,
            text_color=FLOW_TEXT_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))

        if self._model.warnings:
            ctk.CTkLabel(
                head,
                text="⚠ " + " / ".join(self._model.warnings),
                font=SMALL_FONT,
                text_color=FLOW_DANGER,
                anchor="w",
                wraplength=860,
                justify="left",
            ).grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))

    def _build_canvas(self) -> None:
        wrap = ctk.CTkFrame(
            self,
            fg_color=FLOW_SURFACE,
            border_width=CARD_BORDER_WIDTH,
            border_color=FLOW_BORDER,
            corner_radius=CORNER_RADIUS,
        )
        wrap.grid(row=1, column=0, sticky="nsew", padx=12, pady=6)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        self._canvas = tk.Canvas(
            wrap,
            background="#ffffff",
            highlightthickness=0,
            bd=0,
        )
        self._canvas.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        self._canvas.bind("<Configure>", lambda _e: self._draw())

    def _build_footer(self) -> None:
        foot = ctk.CTkFrame(self, fg_color="transparent")
        foot.grid(row=2, column=0, sticky="ew", padx=12, pady=(6, 12))
        foot.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            foot,
            text="高さは概算表示です。Excel 実描画と差が出ることがあります。",
            font=SMALL_FONT,
            text_color=FLOW_TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        btns = ctk.CTkFrame(foot, fg_color="transparent")
        btns.grid(row=0, column=1, sticky="e")

        ctk.CTkButton(
            btns,
            text="キャンセル",
            width=120,
            height=36,
            corner_radius=CORNER_RADIUS,
            fg_color=FLOW_SURFACE,
            hover_color="#f1f5f9",
            text_color=FLOW_TEXT_BODY,
            border_width=CARD_BORDER_WIDTH,
            border_color=FLOW_BORDER,
            command=self._cancel,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btns,
            text="Excelに作成",
            width=140,
            height=36,
            corner_radius=CORNER_RADIUS,
            fg_color=FLOW_ACCENT,
            hover_color=FLOW_ACCENT_HOVER,
            text_color="white",
            font=(FONT_FAMILY, 12, "bold"),
            command=self._confirm,
        ).pack(side="left")

    def _cancel(self) -> None:
        self.grab_release()
        self.destroy()

    def _confirm(self) -> None:
        callback = self._on_confirm
        self.grab_release()
        self.destroy()
        callback()

    def _draw(self) -> None:
        canvas = self._canvas
        canvas.delete("all")
        model = self._model
        if not model.nodes:
            canvas.create_text(
                20,
                20,
                anchor="nw",
                text="表示できるノードがありません",
                fill=FLOW_TEXT_MUTED,
                font=APP_FONT,
            )
            return

        left, top, right, bottom = model.bounds
        content_w = max(1.0, right - left)
        content_h = max(1.0, bottom - top)
        cw = max(1, canvas.winfo_width())
        ch = max(1, canvas.winfo_height())
        scale = min(
            (cw - 2 * PREVIEW_PAD) / content_w,
            (ch - 2 * PREVIEW_PAD) / content_h,
            1.2,
        )
        ox = PREVIEW_PAD - left * scale
        oy = PREVIEW_PAD - top * scale
        by_id = {n.id: n for n in model.nodes}

        for edge in model.edges:
            src = by_id.get(edge.source_id)
            dst = by_id.get(edge.target_id)
            if not src or not dst:
                continue
            x1, y1 = self._anchor(src, edge.direction, scale, ox, oy, start=True)
            x2, y2 = self._anchor(dst, edge.direction, scale, ox, oy, start=False)
            canvas.create_line(
                x1,
                y1,
                x2,
                y2,
                fill=self._connector,
                width=PREVIEW_STROKE,
                arrow=tk.LAST,
            )
            if edge.is_decision:
                label = "Yes" if edge.direction == "down" else "No"
                canvas.create_text(
                    (x1 + x2) / 2 + 8,
                    (y1 + y2) / 2 - 8,
                    text=label,
                    fill=FLOW_TEXT_MUTED,
                    font=(FONT_FAMILY, 9),
                    anchor="w",
                )

        for node in model.nodes:
            self._draw_node(canvas, node, scale, ox, oy)

    def _anchor(
        self,
        node: PlacedNode,
        direction: str,
        scale: float,
        ox: float,
        oy: float,
        *,
        start: bool,
    ) -> tuple[float, float]:
        x = ox + node.left * scale
        y = oy + node.top * scale
        w = node.width * scale
        h = node.height * scale
        if start:
            if direction == "right":
                return x + w, y + h / 2
            return x + w / 2, y + h
        return x + w / 2, y

    def _draw_node(
        self,
        canvas: tk.Canvas,
        node: PlacedNode,
        scale: float,
        ox: float,
        oy: float,
    ) -> None:
        x = ox + node.left * scale
        y = oy + node.top * scale
        w = node.width * scale
        h = node.height * scale
        stroke = self._shape_line
        fill = "#ffffff"

        if node.shape_kind == "diamond":
            canvas.create_polygon(
                [x + w / 2, y, x + w, y + h / 2, x + w / 2, y + h, x, y + h / 2],
                outline=stroke,
                fill=fill,
                width=PREVIEW_STROKE,
            )
        elif node.shape_kind == "parallelogram":
            skew = w * 0.12
            canvas.create_polygon(
                [x + skew, y, x + w, y, x + w - skew, y + h, x, y + h],
                outline=stroke,
                fill=fill,
                width=PREVIEW_STROKE,
            )
        elif node.shape_kind == "manual":
            canvas.create_polygon(
                [x + w * 0.15, y, x + w, y, x + w, y + h, x, y + h],
                outline=stroke,
                fill=fill,
                width=PREVIEW_STROKE,
            )
        else:
            # 端子（角丸）も矩形で近似 — 種別は ID 横とテキストで判別
            canvas.create_rectangle(
                x,
                y,
                x + w,
                y + h,
                outline=stroke,
                fill=fill,
                width=PREVIEW_STROKE,
            )

        label = node.full_text.strip() or f"[{node.id}]"
        canvas.create_text(
            x + w / 2,
            y + h / 2,
            text=label,
            fill=FLOW_TEXT_BODY,
            font=(FONT_FAMILY, max(8, int(10 * min(scale, 1.0)))),
            width=max(20, w - 8),
            justify="center",
        )
        canvas.create_text(
            x + 4,
            y + 2,
            text=node.id,
            fill=FLOW_TEXT_MUTED,
            font=(FONT_FAMILY, 8),
            anchor="nw",
        )
