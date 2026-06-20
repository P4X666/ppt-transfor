"""文本框解析器：TextFrame → Text 模型。

遍历段落与 run，提取字体/颜色/对齐/间距/行距。
对 placeholder 形状，解析继承自 layout/master 的对齐、字号、颜色。
"""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN

from ppt_transfor.models.schema import Font, Paragraph, Run, Text


def _parse_alignment(alignment) -> str | None:
    """对齐方式枚举 → 字符串"""
    if alignment is None:
        return None
    return alignment.name if hasattr(alignment, "name") else str(alignment)


def _parse_anchor(anchor) -> str | None:
    """垂直对齐枚举 → 字符串"""
    if anchor is None:
        return None
    return anchor.name if hasattr(anchor, "name") else str(anchor)


def _parse_auto_size(auto_size) -> str | None:
    """自适应大小枚举 → 字符串。

    NONE（显式无自适应）与 None（继承）统一返回 None，
    避免往返时因 add_textbox 默认行为差异产生噪声。
    """
    if auto_size is None:
        return None
    name = auto_size.name if hasattr(auto_size, "name") else str(auto_size)
    # NONE 与继承统一，往返一致
    if name == "NONE":
        return None
    return name


def _parse_font(font, prs=None, inherited_props: dict | None = None) -> Font:
    """解析 Font 对象。

    对继承的属性（size/color/name 为 None），从 inherited_props 取值固化。

    Args:
        font: python-pptx Font 对象
        prs: Presentation 对象（用于主题色固化）
        inherited_props: placeholder 继承属性 dict
    """
    from ppt_transfor.utils.color import parse_color

    model = Font()

    # 字体名：显式优先，继承兜底
    if font.name is not None:
        model.name = font.name
    elif inherited_props and inherited_props.get("font_name"):
        model.name = inherited_props["font_name"]

    # 字号：显式优先，继承兜底（关键：避免 SHAPE_TO_FIT_TEXT 基于错误字号撑大）
    if font.size is not None:
        model.size = int(font.size)
    elif inherited_props and inherited_props.get("font_size") is not None:
        model.size = inherited_props["font_size"]

    if font.bold is not None:
        model.bold = font.bold
    if font.italic is not None:
        model.italic = font.italic
    if font.underline is not None:
        model.underline = font.underline

    # 颜色：显式优先，继承兜底（主题色已固化为 RGB）
    try:
        color = parse_color(font.color, prs)
        if color is not None:
            model.color = color
        elif inherited_props and inherited_props.get("font_color"):
            model.color = inherited_props["font_color"]
    except Exception:
        if inherited_props and inherited_props.get("font_color"):
            model.color = inherited_props["font_color"]

    return model


def parse_paragraph(para, prs=None, inherited_props: dict | None = None) -> Paragraph:
    """解析单个段落。

    Args:
        para: python-pptx 段落对象
        prs: Presentation 对象（用于主题色固化）
        inherited_props: placeholder 继承属性 dict（对齐/字号/颜色/字体名）
    """
    model = Paragraph()

    # 对齐：显式优先，继承兜底
    align = _parse_alignment(para.alignment)
    if align is not None:
        model.alignment = align
    elif inherited_props and inherited_props.get("alignment"):
        model.alignment = inherited_props["alignment"]

    model.level = para.level

    # 间距：space_before/space_after 是 Length（EMU）
    if para.space_before is not None:
        model.space_before = int(para.space_before)
    if para.space_after is not None:
        model.space_after = int(para.space_after)

    # 行距：float 为倍数，Length 为精确值（存 EMU 整数）
    if para.line_spacing is not None:
        ls = para.line_spacing
        if isinstance(ls, float):
            model.line_spacing = ls
        else:
            # Length 对象，存 EMU 整数
            model.line_spacing = float(int(ls))

    # runs
    for run in para.runs:
        model.runs.append(Run(text=run.text, font=_parse_font(run.font, prs, inherited_props)))

    return model


def parse_text_frame(tf, prs=None, inherited_props: dict | None = None) -> Text:
    """解析 TextFrame → Text 模型。

    Args:
        tf: python-pptx TextFrame 对象
        prs: Presentation 对象（用于主题色固化）
        inherited_props: placeholder 继承属性 dict（对齐/字号/颜色/字体名）
    """
    model = Text()

    # 自动换行
    if tf.word_wrap is not None:
        model.word_wrap = tf.word_wrap

    # 自适应大小
    try:
        if tf.auto_size is not None:
            model.auto_size = _parse_auto_size(tf.auto_size)
    except Exception:
        pass

    # 垂直对齐
    try:
        if tf.vertical_anchor is not None:
            model.vertical_anchor = _parse_anchor(tf.vertical_anchor)
    except Exception:
        pass

    # 段落
    for para in tf.paragraphs:
        model.paragraphs.append(parse_paragraph(para, prs, inherited_props))

    return model
