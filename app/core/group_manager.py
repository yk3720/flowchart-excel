"""MZ0000_FlowchartTool グループ化管理モジュール。

Powered by Auto (Cursor) (rev014)
Rule 2.1.5: 認知負荷の管理に基づき、excel_engine.pyからグループ化ロジックを分離。
"""
import logging
from typing import Any, Dict, List, Tuple
import pywintypes
from app.constants import ExcelConstants, FONT_FAMILY
from app.core.shape_placer import set_text_style

logger = logging.getLogger("flowchart-excel")


def finalize_composites(sheet: Any, d_info: List[Dict], w_fix: float) -> List[Tuple]:
    """判断図形の内部にテキスト枠を重ねる。
    
    Args:
        sheet (Any): Excelワークシートオブジェクト。
        d_info (List[Dict]): 判断図形の情報リスト（shp, l, t, h, txtを含む）。
        w_fix (float): 図形の固定幅。
        
    Returns:
        List[Tuple]: (背景図形, テキスト図形) のペアリスト。
    """
    pairs = []
    for info in d_info:
        txt_shp = sheet.Shapes.AddShape(1, info["l"] + w_fix * 0.125, info["t"], w_fix * 0.75, info["h"])
        txt_shp.Fill.Visible = False
        txt_shp.Line.Visible = False
        set_text_style(txt_shp, info["txt"])
        pairs.append((info["shp"], txt_shp))
    return pairs


def add_frame_and_title(sheet: Any, bounds: Tuple, title: str) -> List[str]:
    """フロー全体を囲む外枠とタイトルを追加する。
    
    Args:
        sheet (Any): Excelワークシートオブジェクト。
        bounds (Tuple): (left, top, right, bottom) の境界座標タプル。
        title (str): タイトルテキスト。
        
    Returns:
        List[str]: 作成したタイトルと外枠の名前リスト。
    """
    names = []
    l, t, r, b = bounds
    w, h = r - l, b - t
    
    # タイトル
    tl, tt = l + (w / 2.0) - 225.0, max(10.0, t - 75.0)
    t_shp = sheet.Shapes.AddTextbox(1, tl, tt, 450.0, 45.0)
    t_shp.Fill.Visible = False
    t_shp.Line.Visible = False
    tf = t_shp.TextFrame2
    tf.TextRange.Text = title
    tf.TextRange.Font.Size = 15
    tf.TextRange.Font.Bold = True
    tf.TextRange.Font.Name = FONT_FAMILY
    tf.TextRange.Font.Fill.ForeColor.RGB = 0
    tf.TextRange.ParagraphFormat.Alignment = ExcelConstants.MSO_ALIGN_CENTER
    tf.VerticalAnchor = ExcelConstants.MSO_ANCHOR_MIDDLE
    names.append(t_shp.Name)
    
    # 外枠
    margin = 30.0
    f_top = min(t, tt)
    f_h = max(b, tt + 45.0) - f_top
    frame = sheet.Shapes.AddShape(1, l - margin, f_top - margin, w + 2 * margin, f_h + 2 * margin)
    frame.Fill.Visible = False
    frame.Line.ForeColor.RGB = 0
    frame.Line.Weight = 1.5
    names.append(frame.Name)
    
    return names


def create_final_groups(sheet: Any, all_names: List[str], composite_pairs: List[Tuple]) -> str:
    """全ての要素を階層的にグループ化する。
    
    Args:
        sheet (Any): Excelワークシートオブジェクト。
        all_names (List[str]): グループ化対象の図形名リスト。
        composite_pairs (List[Tuple]): (背景図形, テキスト図形) のペアリスト。
        
    Returns:
        str: 最終グループ図形の名前。失敗時は空文字列。
    """
    final_names = list(all_names)
    for bg, tx in composite_pairs:
        try:
            grp = sheet.Shapes.Range(tuple([bg.Name, tx.Name])).Group()
            final_names.append(grp.Name)
        except (pywintypes.com_error, AttributeError) as e:
            logger.warning(f"composite_grouping_failed | error={e}")
            final_names.extend([bg.Name, tx.Name])
    
    if not final_names: 
        return ""
    
    try:
        final_grp = sheet.Shapes.Range(tuple(final_names)).Group()
        return final_grp.Name
    except (pywintypes.com_error, AttributeError) as e:
        logger.error(f"grouping_failed | error={e}")
        return ""
