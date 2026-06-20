"""继承解析工具：解析 PPT 中继承自 layout/master/theme 的属性。

PPT 属性继承链：slide → layout → master → theme。
python-pptx 只暴露显式设置的属性，继承属性返回 None。
本模块沿继承链向上查找，把继承值"物化"进 JSON，保证往返保真。

核心能力：
1. 背景继承解析：slide → layout → master
2. placeholder 继承解析：对齐/字号/颜色/字体名
3. 主题色固化：schemeClr → RGB（通过 theme part 的 clrScheme）
"""

from __future__ import annotations

from typing import Optional

from lxml import etree

from ppt_transfor.models.schema import Background, Color
from ppt_transfor.utils.color import parse_color

# DrawingML 命名空间
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _get_fill_from_background(bg_obj, prs=None) -> Optional[tuple[str, Optional[Color]]]:
    """从 background 对象提取填充类型与颜色。

    返回 (type_name, color) 或 None。
    type_name 为 "SOLID"/"BACKGROUND"/"GRADIENT" 等。

    Args:
        bg_obj: python-pptx background 对象
        prs: 所属 Presentation 对象（用于主题色固化，可选）
    """
    try:
        fill = bg_obj.fill
        fill_type = fill.type
        if fill_type is None:
            return None
        type_name = fill_type.name if hasattr(fill_type, "name") else str(fill_type)
        color = None
        if type_name == "SOLID":
            try:
                # 传入 prs 以固化主题色（如 BACKGROUND_1→黑）为 RGB
                color = parse_color(fill.fore_color, prs)
            except Exception:
                pass
        return (type_name, color)
    except Exception:
        return None


def resolve_background(slide, prs=None) -> Optional[Background]:
    """解析幻灯片背景，沿继承链向上查找。

    顺序：slide.background → slide.slide_layout.background → slide.slide_layout.slide_master.background
    取第一个 SOLID 填充作为真实背景。

    Args:
        slide: python-pptx Slide 对象
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        Background 模型，若全链均无显式背景则返回 None
    """
    # 候选背景对象列表（继承链）
    candidates = []
    try:
        candidates.append(slide.background)
    except Exception:
        pass
    try:
        candidates.append(slide.slide_layout.background)
    except Exception:
        pass
    try:
        candidates.append(slide.slide_layout.slide_master.background)
    except Exception:
        pass

    for bg_obj in candidates:
        result = _get_fill_from_background(bg_obj, prs)
        if result is None:
            continue
        type_name, color = result
        # BACKGROUND 类型表示"跟随上层"，继续向上查
        if type_name == "BACKGROUND":
            continue
        # SOLID 类型取真实颜色
        if type_name == "SOLID":
            model = Background(type="solid", color=color)
            return model
        # 其他类型（gradient/pattern 等）记录类型，颜色可能为 None
        model = Background(type=type_name.lower(), color=color)
        return model

    return None


def _get_layout_placeholder(shape, slide):
    """通过 placeholder 索引匹配 layout 中的 placeholder。

    Args:
        shape: placeholder 形状
        slide: 所属幻灯片

    Returns:
        layout 中的 placeholder 对象，找不到返回 None
    """
    try:
        idx = shape.placeholder_format.idx
    except Exception:
        return None

    try:
        layout = slide.slide_layout
    except Exception:
        return None

    # 遍历 layout placeholders 查找匹配索引
    for ph in layout.placeholders:
        try:
            if ph.placeholder_format.idx == idx:
                return ph
        except Exception:
            continue
    return None


def _get_master_placeholder(shape, slide):
    """通过 placeholder 索引匹配 master 中的 placeholder。"""
    try:
        idx = shape.placeholder_format.idx
    except Exception:
        return None

    try:
        master = slide.slide_layout.slide_master
    except Exception:
        return None

    for ph in master.placeholders:
        try:
            if ph.placeholder_format.idx == idx:
                return ph
        except Exception:
            continue
    return None


def _extract_alignment_from_para(para) -> Optional[str]:
    """从段落对象提取对齐方式。"""
    try:
        alignment = para.alignment
        if alignment is None:
            return None
        return alignment.name if hasattr(alignment, "name") else str(alignment)
    except Exception:
        return None


def _extract_font_props(font, prs=None) -> dict:
    """从 Font 对象提取字号、颜色、字体名。

    返回 dict: { "font_size": int|None, "font_color": Color|None, "font_name": str|None }

    Args:
        font: python-pptx Font 对象
        prs: 所属 Presentation 对象（用于主题色固化，可选）
    """
    props = {"font_size": None, "font_color": None, "font_name": None}
    try:
        if font.size is not None:
            props["font_size"] = int(font.size)
    except Exception:
        pass
    try:
        # 传入 prs 以固化主题色（如 TEXT_2→浅灰）为 RGB
        props["font_color"] = parse_color(font.color, prs)
    except Exception:
        pass
    try:
        if font.name is not None:
            props["font_name"] = font.name
    except Exception:
        pass
    return props


def _resolve_schemeclr_to_rgb(scheme_val: str, prs) -> Optional[str]:
    """从 schemeClr 的 val（如 'dk1', 'lt2', 'accent1'）直接解析为 RGB。

    schemeClr 的 val 直接对应 clrScheme 的子元素名，无需经过 MSO_THEME_COLOR 映射。

    Args:
        scheme_val: schemeClr 的 val 属性值（如 "dk1"）
        prs: 所属 Presentation 对象

    Returns:
        RGB 字符串（如 "1F1F1F"），找不到返回 None
    """
    if prs is None:
        return None
    theme_element = _get_theme_element(prs)
    if theme_element is None:
        return None
    clr_scheme = theme_element.find(f".//{{{NS_A}}}clrScheme")
    if clr_scheme is None:
        return None
    color_elem = clr_scheme.find(f"{{{NS_A}}}{scheme_val}")
    if color_elem is None:
        return None
    # 优先 srgbClr，其次 sysClr 的 lastClr
    srgb = color_elem.find(f"{{{NS_A}}}srgbClr")
    if srgb is not None:
        return srgb.get("val")
    sys_clr = color_elem.find(f"{{{NS_A}}}sysClr")
    if sys_clr is not None:
        return sys_clr.get("lastClr") or sys_clr.get("val")
    return None


def _extract_defRPr_from_xml(text_frame_element, prs=None) -> dict:
    """从 txBody 的 lstStyle/lvl1pPr/defRPr 提取默认字体属性。

    layout/master placeholder 的字号、颜色、字体名通常定义在
    <a:lstStyle>/<a:lvl1pPr>/<a:defRPr> 中，python-pptx 不暴露，
    需直接读 XML。

    Args:
        text_frame_element: placeholder 的 _element（<p:sp> 元素）
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        dict: { "font_size": int|None, "font_color": Color|None, "font_name": str|None }
    """
    props = {"font_size": None, "font_color": None, "font_name": None}
    if text_frame_element is None:
        return props

    try:
        # 查找 <p:txBody>/<a:lstStyle>/<a:lvl1pPr>/<a:defRPr>
        tx_body = text_frame_element.find(f".//{{{NS_P}}}txBody")
        if tx_body is None:
            # 兼容直接传入 txBody 的情况
            tx_body = text_frame_element

        lst_style = tx_body.find(f"{{{NS_A}}}lstStyle")
        if lst_style is None:
            return props

        lvl1_pPr = lst_style.find(f"{{{NS_A}}}lvl1pPr")
        if lvl1_pPr is None:
            return props

        def_rPr = lvl1_pPr.find(f"{{{NS_A}}}defRPr")
        if def_rPr is None:
            return props

        # 字号：sz 属性（单位 1/100 pt，如 6000 = 60pt）
        sz = def_rPr.get("sz")
        if sz is not None:
            try:
                # sz 是 1/100 pt，转 EMU：1pt = 12700 EMU
                props["font_size"] = int(int(sz) * 127)
            except (ValueError, TypeError):
                pass

        # 颜色：<a:solidFill>/<a:srgbClr val="..."/>
        solid_fill = def_rPr.find(f"{{{NS_A}}}solidFill")
        if solid_fill is not None:
            srgb = solid_fill.find(f"{{{NS_A}}}srgbClr")
            if srgb is not None:
                val = srgb.get("val")
                if val:
                    props["font_color"] = Color(type="rgb", value=val)
            else:
                # 可能是 schemeClr（主题色），固化为 RGB
                scheme = solid_fill.find(f"{{{NS_A}}}schemeClr")
                if scheme is not None:
                    val = scheme.get("val")
                    if val:
                        # 优先固化主题色为 RGB（如 dk2→浅灰）
                        rgb_value = _resolve_schemeclr_to_rgb(val, prs)
                        if rgb_value is not None:
                            props["font_color"] = Color(type="rgb", value=rgb_value)
                        else:
                            # 固化失败降级保留 theme 类型
                            props["font_color"] = Color(type="theme", value=val)

        # 字体名：<a:latin typeface="..."/>
        latin = def_rPr.find(f"{{{NS_A}}}latin")
        if latin is not None:
            typeface = latin.get("typeface")
            if typeface:
                props["font_name"] = typeface
    except Exception:
        pass

    return props


def _extract_alignment_from_xml(text_frame_element) -> Optional[str]:
    """从 txBody 的段落 pPr 提取对齐方式。

    layout/master placeholder 的对齐通常定义在
    <a:lstStyle>/<a:lvl1pPr>/<a:pPr algn="ctr"/> 中。

    Args:
        text_frame_element: placeholder 的 _element（<p:sp> 元素）

    Returns:
        对齐方式字符串（如 "CENTER"）或 None
    """
    if text_frame_element is None:
        return None

    try:
        tx_body = text_frame_element.find(f".//{{{NS_P}}}txBody")
        if tx_body is None:
            tx_body = text_frame_element

        lst_style = tx_body.find(f"{{{NS_A}}}lstStyle")
        if lst_style is None:
            return None

        lvl1_pPr = lst_style.find(f"{{{NS_A}}}lvl1pPr")
        if lvl1_pPr is None:
            return None

        pPr = lvl1_pPr.find(f"{{{NS_A}}}pPr")
        if pPr is None:
            return None

        algn = pPr.get("algn")
        if algn:
            # algn 值如 "ctr"/"l"/"r"/"just"，映射到 PP_ALIGN 名称
            algn_map = {
                "l": "LEFT",
                "ctr": "CENTER",
                "r": "RIGHT",
                "just": "JUSTIFY",
                "dist": "DISTRIBUTE",
                "thaiDist": "THAI_DISTRIBUTE",
                "justLow": "JUSTIFY_LOW",
            }
            return algn_map.get(algn, algn.upper())
    except Exception:
        pass

    return None


def resolve_placeholder_props(shape, slide, prs=None) -> dict:
    """解析 placeholder 形状的继承属性（对齐、字体、字号、颜色）。

    通过 shape.placeholder_format.idx 匹配 layout/master 的 placeholder，
    读取其 lstStyle/lvl1pPr 中的 defRPr（默认字体）和 pPr（对齐）。
    python-pptx 不暴露这些属性，需直接读 XML。

    Args:
        shape: placeholder 形状
        slide: 所属幻灯片
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        dict: {
            "alignment": str|None,
            "font_size": int|None,
            "font_color": Color|None,
            "font_name": str|None
        }
    """
    result = {
        "alignment": None,
        "font_size": None,
        "font_color": None,
        "font_name": None,
    }

    # 依次尝试 layout placeholder → master placeholder
    for getter in (_get_layout_placeholder, _get_master_placeholder):
        ref_placeholder = getter(shape, slide)
        if ref_placeholder is None:
            continue

        # 优先从 XML 的 lstStyle/lvl1pPr 提取（python-pptx 不暴露 defRPr）
        try:
            elem = ref_placeholder._element
        except Exception:
            elem = None

        if elem is not None:
            # 对齐
            if result["alignment"] is None:
                align = _extract_alignment_from_xml(elem)
                if align is not None:
                    result["alignment"] = align

            # 字体属性（字号/颜色/字体名），传入 prs 固化主题色
            font_props = _extract_defRPr_from_xml(elem, prs)
            for key in ("font_size", "font_color", "font_name"):
                if result[key] is None and font_props.get(key) is not None:
                    result[key] = font_props[key]

        # 兜底：若 XML 未取到，尝试 python-pptx 的 run font
        if not ref_placeholder.has_text_frame:
            continue
        tf = ref_placeholder.text_frame
        if not tf.paragraphs:
            continue

        para = tf.paragraphs[0]
        # 对齐
        if result["alignment"] is None:
            align = _extract_alignment_from_para(para)
            if align is not None:
                result["alignment"] = align

        # 字体属性（取第一个 run），传入 prs 固化主题色
        if para.runs:
            font_props = _extract_font_props(para.runs[0].font, prs)
            for key in ("font_size", "font_color", "font_name"):
                if result[key] is None and font_props.get(key) is not None:
                    result[key] = font_props[key]

        # 若已取到全部属性则提前结束
        if all(v is not None for v in result.values()):
            break

    return result


def _parse_theme_color_rgb(theme_element, scheme_color_name: str) -> Optional[str]:
    """从 theme XML 解析主题色为 RGB 字符串。

    theme 的 clrScheme 结构：
    <a:clrScheme name="...">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F1F1F"/></a:dk2>
      <a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>
      <a:accent1><a:srgbClr val="..."/></a:accent1>
      ...
    </a:clrScheme>

    MSO_THEME_COLOR 名称到 clrScheme 子元素名的映射：
    - ACCENT_1..6 → accent1..6
    - TEXT_1 (dk1), TEXT_2 (dk2), BACKGROUND_1 (lt1), BACKGROUND_2 (lt2)
    - HYPERLINK → hlink
    - FOLLOWED_HYPERLINK → folHlink

    Args:
        theme_element: theme part 的根 XML 元素
        scheme_color_name: MSO_THEME_COLOR 的名称（如 "ACCENT_1"）

    Returns:
        RGB 字符串（如 "FF0000"），找不到返回 None
    """
    if theme_element is None:
        return None

    # MSO_THEME_COLOR 名称 → clrScheme 子元素名映射
    name_map = {
        "TEXT_1": "dk1",
        "TEXT_2": "dk2",
        "BACKGROUND_1": "lt1",
        "BACKGROUND_2": "lt2",
        "ACCENT_1": "accent1",
        "ACCENT_2": "accent2",
        "ACCENT_3": "accent3",
        "ACCENT_4": "accent4",
        "ACCENT_5": "accent5",
        "ACCENT_6": "accent6",
        "HYPERLINK": "hlink",
        "FOLLOWED_HYPERLINK": "folHlink",
        # 兼容无下划线写法
        "ACCENT1": "accent1",
        "ACCENT2": "accent2",
        "ACCENT3": "accent3",
        "ACCENT4": "accent4",
        "ACCENT5": "accent5",
        "ACCENT6": "accent6",
        "TEXT1": "dk1",
        "TEXT2": "dk2",
        "BACKGROUND1": "lt1",
        "BACKGROUND2": "lt2",
    }

    clr_element_name = name_map.get(scheme_color_name)
    if clr_element_name is None:
        # 尝试小写直接映射
        clr_element_name = scheme_color_name.lower()

    # 查找 clrScheme
    clr_scheme = theme_element.find(f".//{{{NS_A}}}clrScheme")
    if clr_scheme is None:
        return None

    # 查找对应颜色元素
    color_elem = clr_scheme.find(f"{{{NS_A}}}{clr_element_name}")
    if color_elem is None:
        return None

    # 颜色元素下可能有 srgbClr（直接 RGB）或 sysClr（系统颜色）
    srgb = color_elem.find(f"{{{NS_A}}}srgbClr")
    if srgb is not None:
        val = srgb.get("val")
        return val

    sys_clr = color_elem.find(f"{{{NS_A}}}sysClr")
    if sys_clr is not None:
        # sysClr 的 lastClr 属性通常是实际 RGB
        val = sys_clr.get("lastClr")
        if val is None:
            val = sys_clr.get("val")
        return val

    return None


def _get_theme_element(presentation) -> Optional[object]:
    """获取 presentation 的 theme XML 元素。

    theme part 挂在 slide_master 上，通过 rels 查找 theme 关系获取。
    python-pptx 的 SlideMasterPart 没有 theme_part 属性，需遍历 rels。
    """
    try:
        master = presentation.slide_masters[0]
        # 通过 rels 查找 theme 关系
        for rel in master.part.rels.values():
            if "theme" in rel.reltype.lower():
                theme_part = rel.target_part
                theme_element = etree.fromstring(theme_part.blob)
                return theme_element
        return None
    except Exception:
        return None


def resolve_theme_color(color_format, presentation) -> Optional[Color]:
    """将主题色解析为具体 RGB，固化为 Color(type="rgb")。

    Args:
        color_format: python-pptx ColorFormat 对象（type 为 SCHEME）
        presentation: 所属 Presentation 对象（用于访问 theme part）

    Returns:
        Color(type="rgb", value="RRGGBB")，解析失败返回 None
    """
    try:
        theme_color = color_format.theme_color
        scheme_name = theme_color.name if hasattr(theme_color, "name") else str(theme_color)
    except Exception:
        return None

    theme_element = _get_theme_element(presentation)
    if theme_element is None:
        return None

    rgb_value = _parse_theme_color_rgb(theme_element, scheme_name)
    if rgb_value is None:
        return None

    return Color(type="rgb", value=rgb_value)
