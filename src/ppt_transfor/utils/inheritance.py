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

import colorsys
from typing import Optional

from lxml import etree

from ppt_transfor.models.schema import Background, Color
from ppt_transfor.utils.color import parse_color

# DrawingML 命名空间
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _apply_color_modifiers(rgb_hex: str, schemeclr_element) -> str:
    """对 RGB 应用 schemeClr 的颜色修饰符（lumMod/lumOff/tint/shade 等）。

    OOXML 中 schemeClr 可携带子元素修饰基础颜色，常见修饰符：
    - lumMod: 亮度 *= val/100000（HSL 空间）
    - lumOff: 亮度 += val/100000（HSL 空间）
    - tint: 颜色推向白色（RGB 线性插值）
    - shade: 颜色推向黑色（RGB 线性插值）
    - satMod/satOff: 饱和度调制/偏移（HSL 空间）
    - hueMod/hueOff: 色相调制/偏移（HSL 空间）
    - comp: 互补色（H + 180°）
    - inv: 反色
    - gray: 灰度化

    HSL 空间修饰符批量应用，避免多次 RGB↔HSL 转换误差累积。

    Args:
        rgb_hex: 基础 RGB 字符串（如 "B3B3BF"）
        schemeclr_element: <a:schemeClr> lxml 元素，其子元素为修饰符

    Returns:
        应用修饰符后的 RGB 字符串，无修饰符或失败时返回原值
    """
    if schemeclr_element is None or not rgb_hex:
        return rgb_hex

    try:
        r = int(rgb_hex[0:2], 16) / 255.0
        g = int(rgb_hex[2:4], 16) / 255.0
        b = int(rgb_hex[4:6], 16) / 255.0
    except (ValueError, IndexError):
        return rgb_hex

    # HSL 空间修饰符累积值（批量应用，避免多次转换）
    lum_mod = 1.0
    lum_off = 0.0
    sat_mod = 1.0
    sat_off = 0.0
    hue_mod = 1.0
    hue_off = 0.0
    has_hsl_mod = False

    for child in schemeclr_element:
        try:
            tag = etree.QName(child).localname
        except Exception:
            continue
        val = child.get("val")
        if val is None:
            continue
        try:
            v = int(val) / 100000.0
        except ValueError:
            continue

        if tag == "tint":
            # tint: 把颜色推向白色
            r = r * (1 - v) + 1.0 * v
            g = g * (1 - v) + 1.0 * v
            b = b * (1 - v) + 1.0 * v
        elif tag == "shade":
            # shade: 把颜色推向黑色
            r = r * v
            g = g * v
            b = b * v
        elif tag == "comp":
            # 互补色：H + 180°
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            h = (h + 0.5) % 1.0
            r, g, b = colorsys.hls_to_rgb(h, l, s)
        elif tag == "inv":
            r = 1.0 - r
            g = 1.0 - g
            b = 1.0 - b
        elif tag == "gray":
            gray = 0.21 * r + 0.72 * g + 0.07 * b
            r = g = b = gray
        elif tag == "lumMod":
            lum_mod *= v
            has_hsl_mod = True
        elif tag == "lumOff":
            lum_off += v
            has_hsl_mod = True
        elif tag == "satMod":
            sat_mod *= v
            has_hsl_mod = True
        elif tag == "satOff":
            sat_off += v
            has_hsl_mod = True
        elif tag == "hueMod":
            hue_mod *= v
            has_hsl_mod = True
        elif tag == "hueOff":
            hue_off += v
            has_hsl_mod = True

    # 批量应用 HSL 空间修饰符
    if has_hsl_mod:
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        h = (h * hue_mod + hue_off) % 1.0
        l = max(0.0, min(1.0, l * lum_mod + lum_off))
        s = max(0.0, min(1.0, s * sat_mod + sat_off))
        r, g, b = colorsys.hls_to_rgb(h, l, s)

    r = max(0, min(255, round(r * 255)))
    g = max(0, min(255, round(g * 255)))
    b = max(0, min(255, round(b * 255)))
    return f"{r:02X}{g:02X}{b:02X}"


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


def _get_background_fill_from_xml(slide_element, prs=None) -> Optional[tuple[str, Optional[Color]]]:
    """从 <p:cSld>/<p:bg>/<p:bgPr> 直接解析背景填充。

    python-pptx 对 slideLayout 的背景报告不准确（有时报告 BACKGROUND 而实际
    XML 中有 <a:solidFill>），直接从 XML 解析更可靠。同时正确处理 schemeClr
    携带的颜色修饰符（lumOff/lumMod 等），确保如 accent6 + lumOff 13725
    能正确计算为 D9D9DF。

    Args:
        slide_element: <p:sld>/<p:sldLayout>/<p:sldMaster> 元素
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        (type_name, color) 或 None。type_name 为 "SOLID"/"BACKGROUND" 等。
    """
    if slide_element is None:
        return None

    try:
        cSld = slide_element.find(f"{{{NS_P}}}cSld")
        if cSld is None:
            return None
        bg = cSld.find(f"{{{NS_P}}}bg")
        if bg is None:
            # 无 <p:bg> 元素，表示继承上层
            return None

        bgPr = bg.find(f"{{{NS_P}}}bgPr")
        if bgPr is None:
            # 可能有 <p:bgRef>（引用主题背景）
            bgRef = bg.find(f"{{{NS_P}}}bgRef")
            if bgRef is not None:
                # bgRef 引用主题背景，视为继承
                return ("BACKGROUND", None)
            return None

        # 从 bgPr 解析 solidFill
        solid_fill = bgPr.find(f"{{{NS_A}}}solidFill")
        if solid_fill is not None:
            srgb = solid_fill.find(f"{{{NS_A}}}srgbClr")
            if srgb is not None:
                val = srgb.get("val")
                if val:
                    return ("SOLID", Color(type="rgb", value=val))
            else:
                scheme = solid_fill.find(f"{{{NS_A}}}schemeClr")
                if scheme is not None:
                    val = scheme.get("val")
                    if val:
                        # 传入 scheme 元素以应用 lumOff/lumMod 等修饰符
                        rgb_value = _resolve_schemeclr_to_rgb(val, prs, scheme)
                        if rgb_value is not None:
                            return ("SOLID", Color(type="rgb", value=rgb_value))
                        else:
                            return ("SOLID", Color(type="theme", value=val))

        # 渐变填充
        grad_fill = bgPr.find(f"{{{NS_A}}}gradFill")
        if grad_fill is not None:
            return ("GRADIENT", None)

        # 无填充
        no_fill = bgPr.find(f"{{{NS_A}}}noFill")
        if no_fill is not None:
            return ("BACKGROUND", None)

        return None
    except Exception:
        return None


def resolve_background(slide, prs=None) -> Optional[Background]:
    """解析幻灯片背景，沿继承链向上查找。

    顺序：slide → slide_layout → slide_master
    取第一个 SOLID 填充作为真实背景。

    优先从 XML 解析（python-pptx 对 layout 背景报告不准确），API 兜底。
    转换后 PPT 使用 Blank 布局，不继承原始 layout/master 的背景，
    因此需要把继承链上的背景"物化"到 slide 自身，保证视觉保真。

    Args:
        slide: python-pptx Slide 对象
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        Background 模型，若全链均无显式背景则返回 None
    """
    # 候选层级（slide/layout/master），同时获取 element 和 background 对象
    candidates = []
    try:
        candidates.append(slide)
    except Exception:
        pass
    try:
        candidates.append(slide.slide_layout)
    except Exception:
        pass
    try:
        candidates.append(slide.slide_layout.slide_master)
    except Exception:
        pass

    for layer in candidates:
        # 1. 优先从 XML 解析（python-pptx API 对 layout 背景报告不准确）
        try:
            elem = layer._element
        except Exception:
            elem = None
        xml_result = _get_background_fill_from_xml(elem, prs)
        if xml_result is not None:
            type_name, color = xml_result
            if type_name == "BACKGROUND":
                continue
            if type_name == "SOLID":
                return Background(type="solid", color=color)
            return Background(type=type_name.lower(), color=color)

        # 2. XML 无结果时用 python-pptx API 兜底
        try:
            bg_obj = layer.background
        except Exception:
            continue
        api_result = _get_fill_from_background(bg_obj, prs)
        if api_result is None:
            continue
        type_name, color = api_result
        if type_name == "BACKGROUND":
            continue
        if type_name == "SOLID":
            return Background(type="solid", color=color)
        return Background(type=type_name.lower(), color=color)

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


def _resolve_schemeclr_to_rgb(scheme_val: str, prs, schemeclr_element=None) -> Optional[str]:
    """从 schemeClr 的 val（如 'dk1', 'lt2', 'accent1'）直接解析为 RGB。

    schemeClr 的 val 直接对应 clrScheme 的子元素名，无需经过 MSO_THEME_COLOR 映射。
    若传入 schemeclr_element，会应用其携带的颜色修饰符（lumOff/lumMod/tint/shade 等），
    确保如 accent6 + lumOff 13725 能正确计算为调整后的 RGB，而非主题色原值。

    Args:
        scheme_val: schemeClr 的 val 属性值（如 "dk1"）
        prs: 所属 Presentation 对象
        schemeclr_element: <a:schemeClr> lxml 元素（可选），用于读取修饰符

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
        rgb = srgb.get("val")
    else:
        sys_clr = color_elem.find(f"{{{NS_A}}}sysClr")
        if sys_clr is not None:
            rgb = sys_clr.get("lastClr") or sys_clr.get("val")
        else:
            return None

    # 应用颜色修饰符（lumOff/lumMod/tint/shade 等），schemeclr_element 为空时原样返回
    if rgb and schemeclr_element is not None:
        rgb = _apply_color_modifiers(rgb, schemeclr_element)
    return rgb


def _extract_defRPr_from_xml(text_frame_element, prs=None) -> dict:
    """从 txBody 的 lstStyle/lvl1pPr/defRPr 提取默认字体属性。

    layout/master placeholder 的字号、颜色、字体名通常定义在
    <a:lstStyle>/<a:lvl1pPr>/<a:defRPr> 中，python-pptx 不暴露，
    需直接读 XML。

    除 latin 字体外，还解析 ea/cs/sym 字体，确保 PowerPoint 字体回退
    行为与原始一致（缺失 ea/cs/sym 可能导致某些字符用默认字体渲染）。

    Args:
        text_frame_element: placeholder 的 _element（<p:sp> 元素）
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        dict: {
            "font_size": int|None,
            "font_color": Color|None,
            "font_name": str|None,      # latin 字体名
            "font_cap": str|None,
            "font_ea": str|None,        # 东亚字体名
            "font_cs": str|None,        # 复杂脚本字体名
            "font_sym": str|None,       # 符号字体名
        }
    """
    props = {
        "font_size": None,
        "font_color": None,
        "font_name": None,
        "font_cap": None,
        "font_ea": None,
        "font_cs": None,
        "font_sym": None,
    }
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

        # 大小写：cap 属性（all/small）
        cap = def_rPr.get("cap")
        if cap:
            props["font_cap"] = cap.lower()

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
                        # 传入 scheme 元素以应用 lumOff/lumMod 等修饰符
                        rgb_value = _resolve_schemeclr_to_rgb(val, prs, scheme)
                        if rgb_value is not None:
                            props["font_color"] = Color(type="rgb", value=rgb_value)
                        else:
                            # 固化失败降级保留 theme 类型
                            props["font_color"] = Color(type="theme", value=val)

        # 字体名：latin/ea/cs/sym 四种字体（python-pptx 只暴露 latin via font.name）
        # 补全 ea/cs/sym 确保 PowerPoint 字体回退行为与原始一致
        for tag, key in (
            ("latin", "font_name"),
            ("ea", "font_ea"),
            ("cs", "font_cs"),
            ("sym", "font_sym"),
        ):
            elem = def_rPr.find(f"{{{NS_A}}}{tag}")
            if elem is not None:
                typeface = elem.get("typeface")
                if typeface:
                    props[key] = typeface
    except Exception:
        pass

    return props


def extract_txbody_default_props(tx_body_element, prs=None) -> dict:
    """从 <a:txBody>/<a:lstStyle> 提取默认文本属性。

    shape 级别的默认文本样式定义在 txBody 的 lstStyle 中，
    python-pptx 不会自动应用到 run.font，需手动解析并作为继承属性
    传给 parse_paragraph / _parse_font。

    当前先处理 <a:lvl1pPr>（大多数文本框只有一级），后续可按 para.level 扩展。

    Args:
        tx_body_element: <p:txBody> 元素（python-pptx TextFrame 的 _element）
        prs: 所属 Presentation 对象（用于主题色固化）

    Returns:
        dict: {
            "alignment": str|None,
            "font_size": int|None,
            "font_color": Color|None,
            "font_name": str|None,
            "font_cap": str|None,
            "font_ea": str|None,
            "font_cs": str|None,
            "font_sym": str|None,
        }
    """
    props = {
        "alignment": None,
        "font_size": None,
        "font_color": None,
        "font_name": None,
        "font_cap": None,
        "font_ea": None,
        "font_cs": None,
        "font_sym": None,
    }
    if tx_body_element is None:
        return props

    try:
        # 对齐：从 lstStyle/lvl1pPr/pPr 提取
        align = _extract_alignment_from_xml(tx_body_element)
        if align is not None:
            props["alignment"] = align

        # 字体属性（字号/颜色/字体名/大小写/ea/cs/sym）：从 lstStyle/lvl1pPr/defRPr 提取
        font_props = _extract_defRPr_from_xml(tx_body_element, prs)
        for key in (
            "font_size", "font_color", "font_name", "font_cap",
            "font_ea", "font_cs", "font_sym",
        ):
            if font_props.get(key) is not None:
                props[key] = font_props[key]
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

        # 对齐可能直接定义在 lvl1pPr@algn，也可能在子 pPr@algn
        algn = lvl1_pPr.get("algn")
        if not algn:
            pPr = lvl1_pPr.find(f"{{{NS_A}}}pPr")
            if pPr is not None:
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


def _extract_master_txstyle_props(master_element, placeholder_type: str, prs=None) -> dict:
    """从 slideMaster 的 txStyles 提取标题/正文/其他样式属性。

    slideMaster 的 txStyles 包含三个独立样式表：
    - titleStyle: 标题 placeholder 的兜底样式
    - bodyStyle: 正文 placeholder 的兜底样式
    - otherStyle: 其他 placeholder 的兜底样式

    当 placeholder 自身的 lstStyle 和 layout/master placeholder 都没有
    某个属性（如 cap）时，回退到 txStyles。典型场景：master titleStyle
    的 lvl1pPr/defRPr 有 cap="all"，但 layout 的 defRPr 没有显式 cap 属性，
    此时 PowerPoint 会继承 master titleStyle 的 cap="all"（缺失表示继承）。

    Args:
        master_element: <p:sldMaster> 元素
        placeholder_type: placeholder 类型字符串（如 "title"/"body"/"ctrTitle"）
        prs: Presentation 对象（用于主题色固化）

    Returns:
        dict: 字体属性 dict
    """
    props = {
        "font_size": None,
        "font_color": None,
        "font_name": None,
        "font_cap": None,
        "font_ea": None,
        "font_cs": None,
        "font_sym": None,
    }
    if master_element is None:
        return props

    # 根据 placeholder 类型选择 txStyles 子元素
    # title/ctrTitle → titleStyle，body → bodyStyle，其他 → otherStyle
    title_types = {"title", "ctrTitle"}
    body_types = {"body", "obj", "txt"}

    txStyles = master_element.find(f"{{{NS_P}}}txStyles")
    if txStyles is None:
        return props

    if placeholder_type in title_types:
        style_elem = txStyles.find(f"{{{NS_P}}}titleStyle")
    elif placeholder_type in body_types:
        style_elem = txStyles.find(f"{{{NS_P}}}bodyStyle")
    else:
        style_elem = txStyles.find(f"{{{NS_P}}}otherStyle")

    if style_elem is None:
        return props

    # 从 lvl1pPr/defRPr 提取属性（结构与 lstStyle/lvl1pPr/defRPr 相同）
    try:
        lvl1_pPr = style_elem.find(f"{{{NS_A}}}lvl1pPr")
        if lvl1_pPr is None:
            return props
        def_rPr = lvl1_pPr.find(f"{{{NS_A}}}defRPr")
        if def_rPr is None:
            return props

        # 字号
        sz = def_rPr.get("sz")
        if sz is not None:
            try:
                props["font_size"] = int(int(sz) * 127)
            except (ValueError, TypeError):
                pass

        # 大小写
        cap = def_rPr.get("cap")
        if cap:
            props["font_cap"] = cap.lower()

        # 颜色
        solid_fill = def_rPr.find(f"{{{NS_A}}}solidFill")
        if solid_fill is not None:
            srgb = solid_fill.find(f"{{{NS_A}}}srgbClr")
            if srgb is not None:
                val = srgb.get("val")
                if val:
                    props["font_color"] = Color(type="rgb", value=val)
            else:
                scheme = solid_fill.find(f"{{{NS_A}}}schemeClr")
                if scheme is not None:
                    val = scheme.get("val")
                    if val:
                        rgb_value = _resolve_schemeclr_to_rgb(val, prs, scheme)
                        if rgb_value is not None:
                            props["font_color"] = Color(type="rgb", value=rgb_value)
                        else:
                            props["font_color"] = Color(type="theme", value=val)

        # 字体名：latin/ea/cs/sym
        for tag, key in (
            ("latin", "font_name"),
            ("ea", "font_ea"),
            ("cs", "font_cs"),
            ("sym", "font_sym"),
        ):
            elem = def_rPr.find(f"{{{NS_A}}}{tag}")
            if elem is not None:
                typeface = elem.get("typeface")
                if typeface:
                    props[key] = typeface
    except Exception:
        pass

    return props


def resolve_placeholder_props(shape, slide, prs=None) -> dict:
    """解析 placeholder 形状的继承属性（对齐、字体、字号、颜色）。

    通过 shape.placeholder_format.idx 匹配 layout/master 的 placeholder，
    读取其 lstStyle/lvl1pPr 中的 defRPr（默认字体）和 pPr（对齐）。
    python-pptx 不暴露这些属性，需直接读 XML。

    继承链：slide placeholder → layout placeholder → master placeholder
            → master txStyles（titleStyle/bodyStyle/otherStyle）
    当某一层缺失属性时，回退到下一层。典型场景：master titleStyle 的
    defRPr 有 cap="all"，但 layout 的 defRPr 没有显式 cap（缺失表示继承），
    此时需回退到 master txStyles 取 cap。

    Args:
        shape: placeholder 形状
        slide: 所属幻灯片
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        dict: {
            "alignment": str|None,
            "font_size": int|None,
            "font_color": Color|None,
            "font_name": str|None,
            "font_cap": str|None,
            "font_ea": str|None,
            "font_cs": str|None,
            "font_sym": str|None,
        }
    """
    result = {
        "alignment": None,
        "font_size": None,
        "font_color": None,
        "font_name": None,
        "font_cap": None,
        "font_ea": None,
        "font_cs": None,
        "font_sym": None,
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

            # 字体属性（字号/颜色/字体名/cap/ea/cs/sym），传入 prs 固化主题色
            font_props = _extract_defRPr_from_xml(elem, prs)
            for key in (
                "font_size", "font_color", "font_name", "font_cap",
                "font_ea", "font_cs", "font_sym",
            ):
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
        # font_cap/ea/cs/sym 无法从 python-pptx 获取，仅兜底 size/color/name
        if para.runs:
            font_props = _extract_font_props(para.runs[0].font, prs)
            for key in ("font_size", "font_color", "font_name"):
                if result[key] is None and font_props.get(key) is not None:
                    result[key] = font_props[key]

        # 若已取到全部属性则提前结束
        if all(v is not None for v in result.values()):
            break

    # 最后兜底：master txStyles（titleStyle/bodyStyle/otherStyle）
    # 当 layout/master placeholder 的 defRPr 缺失某属性时（如 cap），
    # PowerPoint 会继承 master txStyles 的对应属性。
    # 典型：layout defRPr 无 cap，master titleStyle 有 cap="all" → 取 cap="all"
    if not all(v is not None for v in result.values()):
        try:
            placeholder_type = ""
            ph_fmt = shape.placeholder_format
            if ph_fmt is not None and ph_fmt.type is not None:
                placeholder_type = (
                    ph_fmt.type.name if hasattr(ph_fmt.type, "name") else str(ph_fmt.type)
                ).lower()
            master = slide.slide_layout.slide_master
            master_element = master._element
            txstyle_props = _extract_master_txstyle_props(
                master_element, placeholder_type, prs
            )
            for key in (
                "font_size", "font_color", "font_name", "font_cap",
                "font_ea", "font_cs", "font_sym",
            ):
                if result[key] is None and txstyle_props.get(key) is not None:
                    result[key] = txstyle_props[key]
        except Exception:
            pass

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
    若 master 未找到，回退到 presentation.part.rels 兜底。
    """
    rels_sources = []
    try:
        master = presentation.slide_masters[0]
        rels_sources.append(master.part.rels)
    except Exception:
        pass
    try:
        rels_sources.append(presentation.part.rels)
    except Exception:
        pass

    for rels in rels_sources:
        try:
            for rel in rels.values():
                if "theme" in rel.reltype.lower():
                    theme_part = rel.target_part
                    theme_element = etree.fromstring(theme_part.blob)
                    return theme_element
        except Exception:
            continue
    return None


def _get_schemeclr_element(color_format):
    """从 ColorFormat 获取 schemeClr XML 元素（可能含 lumOff/lumMod 子元素）。

    python-pptx 的 ColorFormat 基于 _xFill 元素构造（CT_SolidColorFillProperties），
    SCHEME 类型时 _xFill 即为 <a:solidFill> 元素，其子元素 <a:schemeClr> 才是
    真正的颜色元素（可能携带 lumOff/lumMod/tint/shade 等修饰符子元素）。

    注意：ColorFormat 没有 _element 属性，必须使用 _xFill。早期实现误用
    color_format._element 会抛出 AttributeError 被静默捕获，导致修饰符
    （如 accent6 + lumOff 13725）无法应用，颜色被固化为主题色原值。

    Args:
        color_format: python-pptx ColorFormat 对象

    Returns:
        <a:schemeClr> lxml 元素，获取失败返回 None
    """
    try:
        # ColorFormat 使用 _xFill（CT_SolidColorFillProperties），而非 _element
        elem = getattr(color_format, "_xFill", None)
    except Exception:
        return None
    if elem is None:
        return None
    # _xFill 通常是 <a:solidFill>，其子元素才是 <a:schemeClr>
    try:
        if etree.QName(elem).localname == "schemeClr":
            return elem
    except Exception:
        pass
    # 兜底：查找子元素
    return elem.find(f"{{{NS_A}}}schemeClr")


def resolve_theme_color(color_format, presentation) -> Optional[Color]:
    """将主题色解析为具体 RGB，固化为 Color(type="rgb")。

    同时应用 schemeClr 携带的颜色修饰符（lumOff/lumMod/tint/shade 等），
    确保如 accent6 + lumOff 13725 能正确计算为调整后的 RGB。

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

    # 应用颜色修饰符（lumOff/lumMod/tint/shade 等）
    schemeclr_elem = _get_schemeclr_element(color_format)
    if schemeclr_elem is not None:
        rgb_value = _apply_color_modifiers(rgb_value, schemeclr_elem)

    return Color(type="rgb", value=rgb_value)
