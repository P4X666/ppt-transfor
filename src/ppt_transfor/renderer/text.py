"""文本框渲染器：Text 模型 → TextFrame。

回写段落/run/字体/颜色/对齐/间距/行距。
"""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu

from ppt_transfor.models.schema import Font, Paragraph, Run, Text
from ppt_transfor.utils.color import apply_color


def _apply_alignment(para, alignment: str | None) -> None:
    """对齐方式字符串 → 枚举。

    alignment 为 None 时不设置，保持 add_shape/add_textbox 默认行为。
    注：add_shape 默认创建 CENTER 对齐段落，add_textbox 默认 LEFT。
    对比器将 None ≡ CENTER 等价处理（python-pptx 固有行为差异）。
    """
    if alignment is None:
        return
    try:
        para.alignment = PP_ALIGN[alignment]
    except (KeyError, ValueError):
        pass


def _apply_anchor(tf, anchor: str | None) -> None:
    """垂直对齐字符串 → 枚举"""
    if anchor is None:
        return
    try:
        tf.vertical_anchor = MSO_ANCHOR[anchor]
    except (KeyError, ValueError):
        pass


def _apply_auto_size(tf, auto_size: str | None) -> None:
    """自适应大小字符串 → 枚举"""
    if auto_size is None:
        return
    try:
        tf.auto_size = MSO_AUTO_SIZE[auto_size]
    except (KeyError, ValueError):
        pass


def _apply_font(font_obj, font: Font) -> None:
    """回写字体属性"""
    if font.name is not None:
        font_obj.name = font.name
    if font.size is not None:
        font_obj.size = Emu(font.size)
    if font.bold is not None:
        font_obj.bold = font.bold
    if font.italic is not None:
        font_obj.italic = font.italic
    if font.underline is not None:
        font_obj.underline = font.underline
    if font.color is not None:
        try:
            apply_color(font_obj.color, font.color)
        except Exception:
            pass


def _render_run(para, run_model: Run) -> None:
    """渲染单个 run"""
    run = para.add_run()
    run.text = run_model.text
    _apply_font(run.font, run_model.font)


def render_paragraph(tf, para_model: Paragraph, first: bool = False) -> None:
    """渲染段落到 TextFrame。

    Args:
        tf: python-pptx TextFrame
        para_model: 段落模型
        first: 是否为第一段（复用 TextFrame 自带的空段落）
    """
    # 复用第一段，避免开头出现空段落
    para = tf.paragraphs[0] if first else tf.add_paragraph()

    _apply_alignment(para, para_model.alignment)
    para.level = para_model.level

    if para_model.space_before is not None:
        para.space_before = Emu(para_model.space_before)
    if para_model.space_after is not None:
        para.space_after = Emu(para_model.space_after)
    if para_model.line_spacing is not None:
        para.line_spacing = para_model.line_spacing

    for run_model in para_model.runs:
        _render_run(para, run_model)


def render_text_frame(tf, text_model: Text) -> None:
    """渲染 Text 模型到 TextFrame"""
    # 自动换行
    if text_model.word_wrap is not None:
        tf.word_wrap = text_model.word_wrap

    # 自适应大小：若模型未指定，显式设为 NONE，避免 add_textbox 默认 SHAPE_TO_FIT_TEXT
    if text_model.auto_size is not None:
        _apply_auto_size(tf, text_model.auto_size)
    else:
        try:
            tf.auto_size = MSO_AUTO_SIZE.NONE
        except Exception:
            pass

    # 垂直对齐
    _apply_anchor(tf, text_model.vertical_anchor)

    # 清空默认段落（避免空段）后渲染
    # python-pptx 的 TextFrame 自带一个空段落，第一段复用它
    for idx, para_model in enumerate(text_model.paragraphs):
        render_paragraph(tf, para_model, first=(idx == 0))
