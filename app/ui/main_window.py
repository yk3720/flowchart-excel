"""MZ0000_フローチャート作成(詳細版) メインUI。

Powered by Auto (Cursor) (rev014)
"""
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import threading
import pythoncom
import pywintypes
import logging
from typing import Dict, Optional, Any
from app.core.excel_engine import ExcelFlowchartEngine, get_excel_app
from app.core.shape_placer import set_text_style
from app.core.connector_manager import add_decision_label
from app.ui.preview_dialog import FlowPreviewDialog
from app.ui.studio_preview import run_studio_preview
from app.constants import (
    FONT_FAMILY, TITLE_FONT, APP_FONT, SMALL_FONT, LABEL_FONT, APP_NAME,
    FLOW_ACCENT, FLOW_ACCENT_HOVER, FLOW_SURFACE, FLOW_SURFACE_MUTED, FLOW_SURFACE_SUBTLE,
    FLOW_BORDER, FLOW_BORDER_STRONG, FLOW_TEXT, FLOW_TEXT_BODY, FLOW_TEXT_MUTED,
    FLOW_DANGER, FLOW_DANGER_HOVER, FLOW_SUCCESS_BG, FLOW_SUCCESS_BORDER, FLOW_SUCCESS_TEXT,
    CARD_BORDER_WIDTH,
    CORNER_RADIUS, THEMES, PRESETS, DEFAULT_BOX_HEIGHT, DEFAULT_BOX_WIDTH,
    DEFAULT_GAP_V,
    DEFAULT_GAP_H,
    ExcelConstants,
    TEMPLATE_DATA,
    REVISION,
    TABLE_HEADERS_10_V2,
)

logger = logging.getLogger("flowchart-excel")

# ボタンの見た目バリアント（shadcn 準拠 · SSOT はここのみ）
STYLE_SECONDARY = dict(
    fg_color=FLOW_SURFACE_SUBTLE, hover_color=FLOW_BORDER, text_color=FLOW_TEXT_BODY,
    border_width=CARD_BORDER_WIDTH, border_color=FLOW_BORDER,
)
STYLE_PRIMARY = dict(fg_color=FLOW_ACCENT, hover_color=FLOW_ACCENT_HOVER, text_color="white")
STYLE_DESTRUCTIVE = dict(fg_color=FLOW_DANGER, hover_color=FLOW_DANGER_HOVER, text_color="white")
STYLE_CARD = dict(fg_color=FLOW_SURFACE, border_width=CARD_BORDER_WIDTH, border_color=FLOW_BORDER)


class FlowchartApp(ctk.CTk):
    """フローチャート作成アプリケーションのメインUIクラス。"""
    
    def __init__(self) -> None:
        ctk.set_appearance_mode("light")
        super().__init__()
        self.title(f"{APP_NAME} {REVISION}")
        self.configure(fg_color=FLOW_SURFACE_MUTED)
        
        # 状態保持
        self.is_mini_mode = False
        self.is_help_mode = False
        self.is_processing = False
        self.stop_event = threading.Event()
        self.engine = ExcelFlowchartEngine(self.stop_event)
        
        # 設定変数
        self.var_height = tk.DoubleVar(value=DEFAULT_BOX_HEIGHT)
        self.var_width = tk.DoubleVar(value=DEFAULT_BOX_WIDTH)
        self.var_gap_v = tk.DoubleVar(value=DEFAULT_GAP_V)
        self.var_gap_h = tk.DoubleVar(value=DEFAULT_GAP_H)
        self.var_theme = tk.StringVar(value="標準（信頼）")
        self.status_text = tk.StringVar(value="Excelを待機中...")
        self.preset_buttons: Dict[str, ctk.CTkButton] = {}
        
        self._setup_window()
        self._create_header()
        
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        self.help_frame = ctk.CTkFrame(self, fg_color="transparent")

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self._create_widgets()
        self._create_help_widgets()
        
        # 監視・ポーリング
        for v in [self.var_height, self.var_width, self.var_gap_v, self.var_gap_h]:
            v.trace_add("write", lambda *a: self._update_preset_highlight())
        self._poll_excel_status()

    def _setup_window(self) -> None:
        """ウィンドウの基本属性設定。"""
        self.geometry("400x850+50+50")
        self.minsize(350, 120)
        self.attributes('-topmost', True)
        self.grid_columnconfigure(0, weight=1)

    def _create_header(self) -> None:
        """共通ヘッダー作成。"""
        f_head = ctk.CTkFrame(self, height=40, corner_radius=0, fg_color=FLOW_SURFACE)
        f_head.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        f_head.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(f_head, text="🏠", width=40, height=30, corner_radius=CORNER_RADIUS,
                      command=lambda: None, **STYLE_SECONDARY).grid(row=0, column=0, padx=(5, 2), pady=5)

        self.btn_help = ctk.CTkButton(f_head, text="📖", width=40, height=30, corner_radius=CORNER_RADIUS,
                                     command=self._toggle_help, **STYLE_SECONDARY)
        self.btn_help.grid(row=0, column=1, padx=2)

        ctk.CTkLabel(f_head, text=APP_NAME, font=TITLE_FONT, text_color=FLOW_TEXT).grid(row=0, column=2, padx=10)

        ctk.CTkButton(f_head, text="⬅", width=40, height=30, corner_radius=CORNER_RADIUS,
                      command=self._snap_left, **STYLE_SECONDARY).grid(row=0, column=3, padx=2)
        ctk.CTkButton(f_head, text="➡", width=40, height=30, corner_radius=CORNER_RADIUS,
                      command=self._snap_right, **STYLE_SECONDARY).grid(row=0, column=4, padx=2)

        self.btn_toggle = ctk.CTkButton(f_head, text="小", width=45, height=30,
                                        corner_radius=CORNER_RADIUS, command=self._toggle_mini_mode,
                                        **STYLE_SECONDARY)
        self.btn_toggle.grid(row=0, column=5, padx=2)

        ctk.CTkButton(f_head, text="閉", width=45, height=30, corner_radius=CORNER_RADIUS,
                      command=self.destroy, **STYLE_SECONDARY).grid(row=0, column=6, padx=(2, 5))

        ctk.CTkFrame(self, height=1, corner_radius=0, fg_color=FLOW_BORDER).grid(
            row=1, column=0, sticky="ew", padx=0, pady=0
        )

    def _create_widgets(self) -> None:
        """メインコンテンツのウィジェット配置。"""
        self.content_frame.grid_columnconfigure(0, weight=1)

        # ステータス表示（カード）
        self.status_card = ctk.CTkFrame(self.content_frame, corner_radius=CORNER_RADIUS, **STYLE_CARD)
        self.status_card.grid(row=0, column=0, padx=0, pady=5, sticky="ew")
        self.status_card.grid_columnconfigure(0, weight=1)
        self.lbl_status = ctk.CTkLabel(self.status_card, textvariable=self.status_text, font=SMALL_FONT,
                                     fg_color="transparent", text_color=FLOW_TEXT_BODY, height=50)
        self.lbl_status.grid(row=0, column=0, padx=10, pady=8, sticky="ew")

        # 雛形作成
        f_temp = ctk.CTkFrame(self.content_frame, corner_radius=CORNER_RADIUS, **STYLE_CARD)
        f_temp.grid(row=1, column=0, padx=0, pady=5, sticky="ew")
        f_temp.grid_columnconfigure((0, 1), weight=1)
        for i, (t, m) in enumerate([("基本：判断無", "simple_no"), ("基本：判断有", "simple_yes"),
                                     ("階層：判断無", "complex_no"), ("階層：判断有", "complex_yes")]):
            ctk.CTkButton(f_temp, text=t, command=lambda x=m: self._create_template(x),
                          font=SMALL_FONT, height=40, corner_radius=CORNER_RADIUS,
                          **STYLE_SECONDARY).grid(row=i//2, column=i%2, padx=8, pady=(8 if i < 2 else 4, 8), sticky="ew")

        # スマート・パレット
        f_pal = ctk.CTkFrame(self.content_frame, corner_radius=CORNER_RADIUS, **STYLE_CARD)
        f_pal.grid(row=2, column=0, padx=0, pady=5, sticky="ew")
        f_pal.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(f_pal, text="🎨 スマート・パレット", font=LABEL_FONT, text_color=FLOW_TEXT
                    ).grid(row=0, column=0, columnspan=2, pady=(10, 5))
        for i, (n, v) in enumerate([("端子", "端子"), ("処理", "処理"), ("判断", "判断"), ("入出力", "入出力"), ("手動入力", "手動入力")]):
            ctk.CTkButton(f_pal, text=n, command=lambda x=v: self._smart_input(x),
                          height=35, corner_radius=CORNER_RADIUS, **STYLE_SECONDARY
                          ).grid(row=(i//2)+1, column=i%2, padx=8, pady=5, sticky="ew")

        # 詳細設定
        f_set = ctk.CTkFrame(self.content_frame, corner_radius=CORNER_RADIUS, **STYLE_CARD)
        f_set.grid(row=3, column=0, padx=0, pady=5, sticky="ew")
        f_set.grid_columnconfigure(1, weight=1)
        for i, (l, v) in enumerate([("高さ:", self.var_height), ("幅:", self.var_width), ("行間:", self.var_gap_v), ("列間:", self.var_gap_h)]):
            ctk.CTkLabel(f_set, text=l, font=APP_FONT, text_color=FLOW_TEXT_BODY
                        ).grid(row=i, column=0, padx=10, pady=(8 if i == 0 else 2, 2), sticky="w")
            ctk.CTkEntry(f_set, textvariable=v, width=75, justify="right", corner_radius=CORNER_RADIUS,
                        fg_color=FLOW_SURFACE, text_color=FLOW_TEXT_BODY,
                        border_width=CARD_BORDER_WIDTH, border_color=FLOW_BORDER
                        ).grid(row=i, column=1, padx=10, pady=(8 if i == 0 else 2, 2), sticky="e")

        ctk.CTkOptionMenu(f_set, values=list(THEMES.keys()), variable=self.var_theme, font=SMALL_FONT,
                         corner_radius=CORNER_RADIUS, fg_color=FLOW_SURFACE_SUBTLE, button_color=FLOW_ACCENT,
                         button_hover_color=FLOW_ACCENT_HOVER, text_color=FLOW_TEXT_BODY,
                         dropdown_fg_color=FLOW_SURFACE, dropdown_text_color=FLOW_TEXT_BODY,
                         ).grid(row=4, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

        # プリセット
        f_preset = ctk.CTkFrame(self.content_frame, corner_radius=CORNER_RADIUS, **STYLE_CARD)
        f_preset.grid(row=4, column=0, padx=0, pady=5, sticky="ew")
        f_preset.grid_columnconfigure((0, 1, 2), weight=1)
        for i, p in enumerate(PRESETS):
            btn = ctk.CTkButton(f_preset, text=p["name"], command=lambda x=p: self._apply_preset(x),
                                width=60, corner_radius=CORNER_RADIUS, **STYLE_SECONDARY)
            btn.grid(row=0, column=i, padx=8, pady=10, sticky="ew")
            self.preset_buttons[p["id"]] = btn

        # メンテナンス
        ctk.CTkButton(self.content_frame, text="↩ Undo (直前削除)", command=self._undo_last,
                      corner_radius=CORNER_RADIUS, **STYLE_SECONDARY).grid(row=5, column=0, padx=0, pady=2, sticky="ew")
        ctk.CTkButton(self.content_frame, text="🗑 図面クリア", command=self._clear_canvas,
                      corner_radius=CORNER_RADIUS, **STYLE_DESTRUCTIVE).grid(row=6, column=0, padx=0, pady=2, sticky="ew")

        # 生成ボタン
        f_gen = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        f_gen.grid(row=7, column=0, padx=0, pady=(15, 5), sticky="ew")
        f_gen.grid_columnconfigure(0, weight=1)

        self.btn_gen_sel = ctk.CTkButton(
            f_gen,
            text="🟦 選択範囲を確認して作成",
            command=self._generate_selection,
            corner_radius=CORNER_RADIUS,
            font=(FONT_FAMILY, 13, "bold"),
            height=45,
            **STYLE_SECONDARY,
        )
        self.btn_gen_sel.grid(row=0, column=0, padx=0, pady=(0, 10), sticky="ew")

        self.btn_gen_full = ctk.CTkButton(
            f_gen,
            text="🚀 表全体を確認して作成",
            command=self._generate_full,
            corner_radius=CORNER_RADIUS,
            font=(FONT_FAMILY, 14, "bold"),
            height=55,
            **STYLE_PRIMARY,
        )
        self.btn_gen_full.grid(row=1, column=0, padx=0, pady=0, sticky="ew")

        # 中止ボタン
        self.btn_cancel = ctk.CTkButton(f_gen, text="✋ 中止", command=self._cancel_generation,
                                       corner_radius=CORNER_RADIUS, font=(FONT_FAMILY, 14, "bold"), height=55,
                                       **STYLE_DESTRUCTIVE)

    def _cancel_generation(self) -> None:
        """進行中の生成処理を中止する。"""
        self.stop_event.set()
        self.status_text.set("🛑 中止しています...")
        self.btn_cancel.configure(state="disabled")

    def _set_processing(self, state: bool) -> None:
        """処理中状態に応じたUIの切り替え。
        
        Args:
            state (bool): 処理中の場合True、処理完了の場合False。
        """
        self.is_processing = state
        st = "disabled" if state else "normal"
        self.btn_gen_sel.configure(state=st)
        self.btn_gen_full.configure(state=st)
        
        if state:
            self.btn_gen_full.grid_remove()
            self.btn_cancel.grid(row=1, column=0, sticky="ew")
            self.btn_cancel.configure(state="normal")
            self.status_text.set("🚀 生成中...")
            self.stop_event.clear()
        else:
            self.btn_cancel.grid_remove()
            self.btn_gen_full.grid(row=1, column=0, sticky="ew")
        self.update()

    def _current_config(self) -> Dict[str, float]:
        return {
            "height": self.var_height.get(),
            "width": self.var_width.get(),
            "gap_v": self.var_gap_v.get(),
            "gap_h": self.var_gap_h.get(),
        }

    def _generate_full(self) -> None:
        """表全体: プレビュー必須 → 確認後に Excel 作成。"""
        self._open_preview_then_draw(True)

    def _generate_selection(self) -> None:
        """選択範囲: プレビュー必須 → 確認後に Excel 作成。"""
        self._open_preview_then_draw(False)

    def _open_preview_then_draw(self, is_full: bool) -> None:
        """プレビュー窓を開き、確定時のみ描画ワーカーを起動する。"""
        if self.is_processing:
            return
        try:
            config = self._current_config()
            payload = self.engine.build_studio_payload(is_full, config)
        except Exception as e:
            logger.exception("preview_failed")
            messagebox.showerror(
                "プレビュー失敗",
                f"【状況】プレビューを作れませんでした。\n"
                f"【原因】{e}\n"
                f"【具体的アクション】Excel を起動し、表を選択してから再試行してください。",
            )
            return

        confirmed = run_studio_preview(payload)
        if confirmed is True:
            self._start_draw_worker(is_full)
            return
        if confirmed is False:
            return

        # dist 未ビルド時は従来 Canvas にフォールバック
        logger.warning("studio_preview_fallback_canvas")
        try:
            model = self.engine.build_preview(is_full, config)
        except Exception as e:
            logger.exception("canvas_preview_failed")
            messagebox.showerror(
                "プレビュー失敗",
                f"studio プレビュー用 dist が無く、Canvas フォールバックも失敗しました。\n{e}\n"
                "preview-web で npm run build を実行してください。",
            )
            return
        if not model.nodes:
            messagebox.showwarning(
                "データなし",
                "有効なノードがありません。ID が数値の行があるか確認してください。",
            )
            return
        theme = THEMES[self.var_theme.get()]
        FlowPreviewDialog(
            self,
            model,
            shape_line_bgr=theme["shape_line"],
            connector_bgr=theme["connector"],
            on_confirm=lambda: self._start_draw_worker(is_full),
        )

    def _start_draw_worker(self, is_full: bool) -> None:
        """プレビュー確定後の Excel 描画を非同期開始する。"""
        threading.Thread(target=self._worker, args=(is_full,), daemon=True).start()

    def _worker(self, is_full: bool) -> None:
        """非同期実行用ワーカー（プレビュー確定後のみ呼ばれる）。

        Args:
            is_full (bool): 表全体から生成する場合True、選択範囲のみの場合False。
        """
        pythoncom.CoInitialize()
        self.after(0, lambda: self._set_processing(True))

        try:
            config = self._current_config()
            theme = THEMES[self.var_theme.get()]

            group_name = self.engine.draw(is_full, config, theme)

            if self.stop_event.is_set():
                self.after(0, lambda: messagebox.showinfo("中止", "描画処理を中止しました。"))
            elif group_name:
                self.engine.last_group_name = group_name
                self.after(
                    0,
                    lambda: messagebox.showinfo("完了", "フローチャートの作成が完了しました。"),
                )

        except Exception as e:
            logger.exception("worker_failed")
            msg = (
                f"【状況】生成処理が中断されました。\n"
                f"【原因】{str(e)}\n"
                f"【具体的アクション】Excelがセル編集中ではないか確認し、編集を終了させてから再試行してください。"
            )
            self.after(0, lambda: messagebox.showerror("エラー", msg))
        finally:
            self.after(0, lambda: self._set_processing(False))
            pythoncom.CoUninitialize()

    def _poll_excel_status(self) -> None:
        """Excelの選択状態を監視してUIに反映。"""
        try:
            app = get_excel_app()
            if not app:
                self.status_text.set("Excel未起動")
            else:
                sel = app.Selection
                r = sel if sel.Count > 1 else sel.CurrentRegion
                title, sheet = "未検出", app.ActiveSheet
                for i in range(-5, 1):
                    row_idx = max(1, r.Cells(1, 1).Row + i)
                    c = sheet.Cells(row_idx, r.Cells(1, 1).Column)
                    if c.Interior.Color == ExcelConstants.TITLE_BG_COLOR and c.Value:
                        title = str(c.Value) if c.Value else "名称未設定"
                        break
                addr = sel.Address.replace('$', '')
                self.status_text.set(f"対象: {title}\n範囲: {addr} ({r.Rows.Count}行)")
        except (AttributeError, pywintypes.com_error):
            if not self.is_processing:
                self.status_text.set("Excel操作中...")
        self.after(1000, self._poll_excel_status)

    def _create_template(self, mode: str) -> None:
        """10列 v2（flowchart-studio 互換）の雛形テーブルを生成。

        Args:
            mode (str): 雛形モード（"simple_no", "simple_yes", "complex_no", "complex_yes"）。
        """
        app = get_excel_app()
        if not app:
            messagebox.showwarning("警告", "Excelを起動してください。")
            return
        try:
            sheet = app.ActiveSheet
            base = app.Selection.Cells(1, 1)
            r, c = base.Row, base.Column
            
            title_map = {"simple_no": "基本：判断無", "simple_yes": "基本：判断有", "complex_no": "階層：判断無", "complex_yes": "階層：判断有"}
            
            cell = sheet.Cells(r, c)
            cell.Value = title_map[mode]
            cell.Font.Bold = True
            cell.Interior.Color = ExcelConstants.TITLE_BG_COLOR
            r += 1
            
            headers = TABLE_HEADERS_10_V2
            for i, h in enumerate(headers):
                target = sheet.Cells(r, c + i)
                target.Value = h
                target.Font.Bold = True
                target.Interior.Color = ExcelConstants.HEADER_BG_COLOR

            data = TEMPLATE_DATA[mode]
            last_col = len(headers) - 1

            for i, row in enumerate(data):
                for j, v in enumerate(row):
                    sheet.Cells(r + 1 + i, c + j).Value = v

            last_cell = sheet.Cells(r + len(data), c + last_col)
            tbl = sheet.ListObjects.Add(1, sheet.Range(sheet.Cells(r, c), last_cell), None, 1)
            tbl.TableStyle = "TableStyleMedium2"

            sheet.Range(
                sheet.Cells(r + 1, c + 3),
                sheet.Cells(r + len(data), c + 4),
            ).HorizontalAlignment = ExcelConstants.XL_RIGHT
            sheet.Range(
                sheet.Cells(r + 1, c + 7),
                sheet.Cells(r + len(data), c + last_col),
            ).ShrinkToFit = True
            
            v_range = sheet.Range(sheet.Cells(r+1, c+1), sheet.Cells(r+100, c+1))
            v_range.Validation.Delete()
            v_range.Validation.Add(3, 1, 1, "端子,処理,判断,入出力,手動入力")
            
            messagebox.showinfo("完了", "雛形を作成しました。")
        except (pywintypes.com_error, AttributeError) as e:
            logger.error(f"template_creation_failed | error={e}")
            messagebox.showerror("エラー", f"作成失敗: {e}")

    def _smart_input(self, stype: str) -> None:
        """スマート・パレットボタンクリックでExcel上に直接図形を生成する。
        
        Args:
            stype (str): 図形種別（"端子", "処理", "判断", "入出力", "手動入力"）。
        """
        app = get_excel_app()
        if not app:
            messagebox.showwarning("警告", "Excelを起動してください。")
            return
        
        try:
            # 1. Excel上で選択セルの位置を取得
            sel = app.Selection
            sheet = app.ActiveSheet
            
            # 2. 選択セルの座標（Left, Top）を基準位置として使用
            cell = sel.Cells(1, 1)
            left_pos = float(cell.Left)
            top_pos = float(cell.Top)
            
            # 3. ボタン種別に応じた図形種別コードを決定
            stype_code = ExcelConstants.MSOSHAPE_RECTANGLE
            is_diamond = False
            is_manual = False
            
            if "判断" in stype:
                stype_code = ExcelConstants.MSOSHAPE_DIAMOND
                is_diamond = True
            elif any(x in stype for x in ["端子", "開始", "終了"]):
                stype_code = ExcelConstants.MSOSHAPE_ROUNDED_RECTANGLE
            elif any(x in stype for x in ["入出力", "データ"]):
                stype_code = ExcelConstants.MSOSHAPE_PARALLELOGRAM
            elif "手動入力" in stype:
                stype_code = ExcelConstants.MSOSHAPE_MANUAL_INPUT
                is_manual = True
            
            # 4. 現在の設定値（var_height, var_width）を使用して図形サイズを決定
            width = self.var_width.get()
            height = self.var_height.get()
            
            # 判断図形の場合は高さを1.3倍
            if is_diamond:
                height = height * 1.3
            
            # 5. 図形を生成
            shp = sheet.Shapes.AddShape(stype_code, left_pos, top_pos, width, height)
            
            # スタイル設定
            shp.Fill.ForeColor.RGB = 0xFFFFFF
            theme = THEMES[self.var_theme.get()]
            shp.Line.ForeColor.RGB = theme["shape_line"]
            
            # テキスト設定（プレースホルダーテキスト「XXXX」）
            set_text_style(shp, "XXXX", is_manual=is_manual)
            
            # 6. コネクタを生成（rev013追加、rev014継承）
            # すべての図形の下側からカギ付き矢印コネクタを生成
            # コネクタの終点を設定するため、一時的な図形を作成して接続（非表示のまま保持）
            temp_bottom_left = left_pos + width / 2  # 図形の下側中央
            temp_bottom_top = top_pos + height + height  # 図形の下側から高さ分下
            temp_bottom = sheet.Shapes.AddShape(ExcelConstants.MSOSHAPE_RECTANGLE, 
                                                temp_bottom_left, temp_bottom_top, 1, 1)
            temp_bottom.Visible = False  # 非表示（コネクタ接続のために保持）
            temp_bottom.Line.Visible = False  # 枠線も非表示
            temp_bottom.Fill.Visible = False  # 塗りつぶしも非表示
            
            try:
                # 下側コネクタを作成
                conn_bottom = sheet.Shapes.AddConnector(ExcelConstants.MSOCONNECTOR_ELBOW, 0, 0, 10, 10)
                conn_bottom.ConnectorFormat.BeginConnect(shp, ExcelConstants.CONNECTOR_SITE_BOTTOM)
                conn_bottom.ConnectorFormat.EndConnect(temp_bottom, ExcelConstants.CONNECTOR_SITE_TOP)
                conn_bottom.Line.ForeColor.RGB = theme["connector"]
                conn_bottom.Line.Weight = 2.25
                conn_bottom.Line.EndArrowheadStyle = 3
                
                # 判断図形の場合は下側コネクタに「Yes」ラベルを配置
                lbl_yes = None
                lbl_no = None
                if is_diamond:
                    lbl_yes = add_decision_label(sheet, shp, "down")
                    if not lbl_yes:
                        logger.warning("decision_label_yes_creation_failed")
                
                # 判断図形の場合は右側からもコネクタを生成
                if is_diamond:
                    temp_right_left = left_pos + width + width  # 図形の右側から幅分右
                    temp_right_top = top_pos + height / 2  # 図形の中央
                    temp_right = sheet.Shapes.AddShape(ExcelConstants.MSOSHAPE_RECTANGLE,
                                                       temp_right_left, temp_right_top, 1, 1)
                    temp_right.Visible = False  # 非表示（コネクタ接続のために保持）
                    temp_right.Line.Visible = False  # 枠線も非表示
                    temp_right.Fill.Visible = False  # 塗りつぶしも非表示
                    
                    try:
                        # 右側コネクタを作成
                        conn_right = sheet.Shapes.AddConnector(ExcelConstants.MSOCONNECTOR_ELBOW, 0, 0, 10, 10)
                        conn_right.ConnectorFormat.BeginConnect(shp, ExcelConstants.CONNECTOR_SITE_RIGHT)
                        conn_right.ConnectorFormat.EndConnect(temp_right, ExcelConstants.CONNECTOR_SITE_LEFT)
                        conn_right.Line.ForeColor.RGB = theme["connector"]
                        conn_right.Line.Weight = 2.25
                        conn_right.Line.EndArrowheadStyle = 3
                        
                        # 判断図形の右側コネクタに「No」ラベルを配置
                        lbl_no = add_decision_label(sheet, shp, "right")
                        if not lbl_no:
                            logger.warning("decision_label_no_creation_failed")
                    except (pywintypes.com_error, AttributeError) as e:
                        logger.error(f"right_connector_creation_failed | error={e}")
                
                # 判断図形の場合は図形とYes/Noラベルをグループ化
                if is_diamond:
                    try:
                        group_names = [shp.Name]
                        if lbl_yes:
                            group_names.append(lbl_yes)
                        if lbl_no:
                            group_names.append(lbl_no)
                        
                        if len(group_names) > 1:
                            decision_group = sheet.Shapes.Range(tuple(group_names)).Group()
                            logger.info(f"decision_group_created | group_name={decision_group.Name} | members={group_names}")
                    except (pywintypes.com_error, AttributeError) as e:
                        logger.error(f"decision_grouping_failed | error={e}")
                
            except (pywintypes.com_error, AttributeError) as e:
                logger.error(f"bottom_connector_creation_failed | error={e}")
            
            logger.info(f"smart_palette_shape_created | type={stype} | position=({left_pos}, {top_pos}) | size=({width}, {height})")
            
        except (pywintypes.com_error, AttributeError) as e:
            logger.error(f"smart_input_failed | error={e}")
            messagebox.showerror("エラー", f"図形の生成に失敗しました: {e}")

    def _undo_last(self) -> None:
        """直前の描画グループを削除する。"""
        if not self.engine.last_group_name: 
            return
        try:
            get_excel_app().ActiveSheet.Shapes(self.engine.last_group_name).Delete()
            self.engine.last_group_name = None
        except (pywintypes.com_error, AttributeError) as e:
            logger.warning(f"undo_failed | error={e}")

    def _clear_canvas(self) -> None:
        """シート上の全図形を削除する。"""
        if not messagebox.askyesno("確認", "全ての図形を削除しますか？"): 
            return
        app = get_excel_app()
        if not app: 
            return
        try:
            for s in list(app.ActiveSheet.Shapes):
                if s.Type in [1, 4, 6]: 
                    s.Delete()
            self.engine.last_group_name = None
        except (pywintypes.com_error, AttributeError) as e:
            logger.warning(f"clear_canvas_failed | error={e}")

    def _apply_preset(self, p: Dict) -> None:
        """プリセット設定を適用する。
        
        Args:
            p (Dict): プリセット設定辞書（height, width, gap_v, gap_hを含む）。
        """
        self.var_height.set(p["height"])
        self.var_width.set(p["width"])
        self.var_gap_v.set(p["gap_v"])
        self.var_gap_h.set(p["gap_h"])

    def _update_preset_highlight(self) -> None:
        """現在の設定値に一致するプリセットボタンをハイライトする。"""
        try:
            h, w, gv, gh = self.var_height.get(), self.var_width.get(), self.var_gap_v.get(), self.var_gap_h.get()
            for p in PRESETS:
                match = all(abs(x - y) < 0.01 for x, y in [(h, p["height"]), (w, p["width"]), (gv, p["gap_v"]), (gh, p["gap_h"])])
                btn = self.preset_buttons[p["id"]]
                if match:
                    btn.configure(fg_color=FLOW_SUCCESS_BG, hover_color=FLOW_SUCCESS_BG,
                                  text_color=FLOW_SUCCESS_TEXT, border_color=FLOW_SUCCESS_BORDER)
                else:
                    btn.configure(**STYLE_SECONDARY)
        except (KeyError, AttributeError) as e:
            logger.warning(f"preset_highlight_update_failed | error={e}")

    def _snap_left(self) -> None:
        """ウィンドウを画面左端に移動する。"""
        self.geometry(f"{self.winfo_width()}x{self.winfo_height()}+0+{self.winfo_y()}")

    def _snap_right(self) -> None:
        """ウィンドウを画面右端に移動する。"""
        self.geometry(f"{self.winfo_width()}x{self.winfo_height()}+{self.winfo_screenwidth() - self.winfo_width()}+{self.winfo_y()}")

    def _toggle_help(self) -> None:
        """ヘルプ画面の表示/非表示を切り替える。"""
        self.is_help_mode = not self.is_help_mode
        if self.is_help_mode:
            self.content_frame.grid_remove()
            self.help_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
            self.btn_help.configure(fg_color=FLOW_ACCENT, hover_color=FLOW_ACCENT_HOVER, text_color="white")
        else:
            self.help_frame.grid_remove()
            self.content_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
            self.btn_help.configure(**STYLE_SECONDARY)

    def _create_help_widgets(self) -> None:
        """ヘルプ画面のウィジェットを作成する。"""
        self.help_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.help_frame, text="📖 操作ガイド", font=LABEL_FONT, text_color=FLOW_TEXT).grid(row=0, column=0, pady=10)
        txt = ctk.CTkTextbox(self.help_frame, font=SMALL_FONT, height=350,
                             fg_color=FLOW_SURFACE, text_color=FLOW_TEXT_BODY,
                             border_width=CARD_BORDER_WIDTH, border_color=FLOW_BORDER, corner_radius=CORNER_RADIUS)
        help_text = (
            "1. Excelでフローチャートにしたい表を選択します。\n"
            "   (ID, 種別, 色, 接続先, 段, 列, Text...)\n\n"
            "2. 「🚀 表全体を確認して作成」を押すと、\n"
            "   flowchart-studio と同じ React Flow プレビューが\n"
            "   開きます。問題なければ「Excelに作成」です。\n\n"
            "3. 初回は preview-web で npm run build が必要です。\n\n"
            "4. 「スマート・パレット」は選択セル位置に\n"
            "   図形を直接置きます（プレビュー対象外）。\n\n"
            "5. 作成中は「中止」で安全に停止できます。"
        )
        txt.insert("0.0", help_text)
        txt.configure(state="disabled")
        txt.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkButton(self.help_frame, text="戻る", command=self._toggle_help, corner_radius=CORNER_RADIUS,
                      **STYLE_SECONDARY).grid(row=2, column=0, pady=10)

    def _toggle_mini_mode(self) -> None:
        """ミニモードの表示/非表示を切り替える。"""
        self.is_mini_mode = not self.is_mini_mode
        if self.is_mini_mode:
            if self.is_help_mode:
                self._toggle_help()
            self.geometry("400x120")
            for child in self.content_frame.winfo_children():
                if child != self.status_card:
                    child.grid_remove()
            self.btn_toggle.configure(text="大")
        else:
            self.geometry("400x850")
            for child in self.content_frame.winfo_children():
                child.grid()
            self.btn_toggle.configure(text="小")
