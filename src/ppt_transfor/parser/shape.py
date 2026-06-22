"""形状解析器：通用属性提取 + 按类型分发。

通用属性：shape_id/name/shape_type/位置/旋转/填充/边框/阴影
类型分发：table/group/picture/connector/autoshape/text_box/placeholder
对 placeholder 形状，解析继承自 layout/master 的对齐、字号、颜色。
"""

from __future__ import annotations

from os.path import normpath
from pathlib import Path

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE

from ppt_transfor.models.schema import Color, Fill, GradientStop, Line, Shadow, Shape
from ppt_transfor.utils.color import parse_color

# XML 命名空间
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _parse_gradient_from_xml(shape_element, prs=None) -> tuple[str | None, float | None, list[GradientStop]]:
    """从 <a:spPr> 下读取渐变填充的 stops、类型和角度。

    Returns:
        (gradient_type, gradient_angle, gradient_stops)
    """
    stops: list[GradientStop] = []
    gradient_type = None
    gradient_angle = None

    if shape_element is None:
        return gradient_type, gradient_angle, stops

    try:
        spPr = shape_element.find(f"{{{NS_P}}}spPr")
        if spPr is None:
            return gradient_type, gradient_angle, stops

        gradFill = spPr.find(f"{{{NS_A}}}gradFill")
        if gradFill is None:
            return gradient_type, gradient_angle, stops

        # 渐变类型与角度
        lin = gradFill.find(f"{{{NS_A}}}lin")
        if lin is not None:
            gradient_type = "linear"
            ang = lin.get("ang")
            if ang is not None:
                try:
                    gradient_angle = float(ang)
                except (ValueError, TypeError):
                    pass

        path = gradFill.find(f"{{{NS_A}}}path")
        if path is not None:
            gradient_type = path.get("path", "path")

        rect = gradFill.find(f"{{{NS_A}}}rect")
        if rect is not None:
            gradient_type = "rect"

        # stops
        gsLst = gradFill.find(f"{{{NS_A}}}gsLst")
        if gsLst is not None:
            for gs in gsLst:
                pos = gs.get("pos")
                if pos is None:
                    continue
                try:
                    position = int(pos) / 100000.0
                except (ValueError, TypeError):
                    continue

                color = None
                # srgbClr：优先直接子元素，其次兼容 <a:solidFill>/<a:srgbClr> 写法
                srgb = gs.find(f"{{{NS_A}}}srgbClr")
                if srgb is None:
                    solid_fill = gs.find(f"{{{NS_A}}}solidFill")
                    if solid_fill is not None:
                        srgb = solid_fill.find(f"{{{NS_A}}}srgbClr")
                if srgb is not None:
                    val = srgb.get("val")
                    if val:
                        color = Color(type="rgb", value=val)

                # schemeClr：固化为 RGB；同样兼容直接子元素或包在 solidFill 中
                if color is None:
                    scheme = gs.find(f"{{{NS_A}}}schemeClr")
                    if scheme is None:
                        solid_fill = gs.find(f"{{{NS_A}}}solidFill")
                        if solid_fill is not None:
                            scheme = solid_fill.find(f"{{{NS_A}}}schemeClr")
                    if scheme is not None:
                        val = scheme.get("val")
                        if val:
                            from ppt_transfor.utils.inheritance import _resolve_schemeclr_to_rgb

                            rgb_value = _resolve_schemeclr_to_rgb(val, prs)
                            if rgb_value is not None:
                                color = Color(type="rgb", value=rgb_value)
                            else:
                                color = Color(type="theme", value=val)

                if color is not None:
                    stops.append(GradientStop(position=position, color=color))
    except Exception:
        pass

    return gradient_type, gradient_angle, stops


def _parse_fill(fill, shape_element=None, prs=None) -> Fill | None:
    """解析填充。

    优先读取 XML：若 <a:spPr> 下显式存在 <a:noFill/>，则视为透明填充。
    python-pptx 对 noFill 形状有时会报告 fill.type=BACKGROUND，
    直接按 XML 判断可避免透明文本框被误判为背景填充。

    None（继承）返回 None，避免往返时因 add_shape 默认填充行为差异产生噪声。
    """
    # 优先检查 XML：显式 noFill 表示透明
    if shape_element is not None:
        try:
            spPr = shape_element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}spPr")
            if spPr is not None:
                no_fill = spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}noFill")
                if no_fill is not None:
                    model = Fill()
                    model.type = "none"
                    return model
        except Exception:
            pass

    try:
        fill_type = fill.type
    except Exception:
        return None

    if fill_type is None:
        return None  # 继承样式，不记录

    # fill_type 是 MSO_FILL 枚举
    type_name = fill_type.name if hasattr(fill_type, "name") else str(fill_type)

    # BACKGROUND 表示“跟随幻灯片背景”，需保留以正确回写
    if type_name == "BACKGROUND":
        model = Fill()
        model.type = "background"
        return model

    model = Fill()
    if type_name == "SOLID":
        try:
            color = parse_color(fill.fore_color, prs)
        except Exception:
            color = None
        if color is not None:
            model.type = "solid"
            model.color = color
        else:
            # 无法解析颜色时降级为无填充，避免渲染为默认黑色
            model.type = "none"
        return model
    elif type_name == "GRADIENT":
        model.type = "gradient"
        grad_type, grad_angle, grad_stops = _parse_gradient_from_xml(shape_element, prs)
        if grad_type is not None:
            model.gradient_type = grad_type
        if grad_angle is not None:
            model.gradient_angle = grad_angle
        if grad_stops:
            model.gradient_stops = grad_stops
    elif type_name == "PATTERNED":
        model.type = "pattern"
    elif type_name == "PICTURE":
        model.type = "picture"
    else:
        model.type = type_name.lower()

    return model


def _parse_line(line, prs=None) -> Line | None:
    """解析边框线条。width=0 视为无线条（add_shape 默认行为）。"""
    model = Line()

    has_value = False
    try:
        if line.width is not None:
            width_val = int(line.width)
            # width=0 视为无线条，不记录
            if width_val > 0:
                model.width = width_val
                has_value = True
    except Exception:
        pass

    try:
        color = parse_color(line.color, prs)
        if color is not None:
            model.color = color
            has_value = True
    except Exception:
        pass

    try:
        dash = line.dash_style
        if dash is not None:
            model.dash = dash.name if hasattr(dash, "name") else str(dash)
            has_value = True
    except Exception:
        pass

    return model if has_value else None


def _parse_shadow(shape) -> Shadow | None:
    """解析阴影。

    inherit=False（显式无阴影）与 None（不记录）统一，
    避免往返时因 add_shape 默认继承阴影产生噪声。
    """
    try:
        shadow = shape.shadow
        if shadow is not None:
            # inherit=False 表示显式关闭阴影，与不记录统一
            if not shadow.inherit:
                return None
            return Shadow(enabled=True)
    except Exception:
        pass
    return None


def _parse_common_props(shape, prs=None) -> dict:
    """解析通用属性（位置/旋转/填充/边框/阴影）"""
    props = {}

    # 位置与尺寸（EMU 整数）
    try:
        props["left"] = int(shape.left) if shape.left is not None else None
    except Exception:
        props["left"] = None
    try:
        props["top"] = int(shape.top) if shape.top is not None else None
    except Exception:
        props["top"] = None
    try:
        props["width"] = int(shape.width) if shape.width is not None else None
    except Exception:
        props["width"] = None
    try:
        props["height"] = int(shape.height) if shape.height is not None else None
    except Exception:
        props["height"] = None

    # 旋转角度（度，浮点）
    try:
        props["rotation"] = float(shape.rotation or 0.0)
    except Exception:
        props["rotation"] = 0.0

    # 填充
    try:
        fill = _parse_fill(shape.fill, getattr(shape, "_element", None), prs)
        if fill is not None:
            props["fill"] = fill
    except Exception:
        pass

    # 边框
    try:
        line = _parse_line(shape.line, prs)
        if line is not None:
            props["line"] = line
    except Exception:
        pass

    # 阴影
    shadow = _parse_shadow(shape)
    if shadow is not None:
        props["shadow"] = shadow

    return props


def _is_placeholder(shape) -> bool:
    """判断形状是否为 placeholder"""
    try:
        return shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER
    except Exception:
        return False


def _placeholder_default_alignment(shape) -> str:
    """根据 placeholder 类型返回默认对齐方式。

    标题/副标题类占位符在 PowerPoint 中默认居中对齐；
    正文/其他占位符默认左对齐。
    """
    try:
        ph_type = shape.placeholder_format.type
        if ph_type is not None:
            type_name = ph_type.name if hasattr(ph_type, "name") else str(ph_type)
            if type_name in ("TITLE", "CENTER_TITLE", "SUBTITLE"):
                return "CENTER"
    except Exception:
        pass
    return "LEFT"


def _parse_chart_part(shape, slide) -> str | None:
    """从 slide part 的 relationships 中定位 chart 对应的 part 路径。

    返回标准化路径，例如 "ppt/charts/chart1.xml"。
    """
    if slide is None:
        return None

    try:
        graphic_data = shape._element.find(f".//{{{NS_A}}}graphicData")
        if graphic_data is None:
            return None

        chart_el = graphic_data.find(f"{{{NS_C}}}chart")
        if chart_el is None:
            return None

        chart_rid = chart_el.get(f"{{{NS_R}}}id")
        if not chart_rid:
            return None

        rel = slide.part.rels.get(chart_rid)
        if rel is None:
            return None

        target = rel.target_ref
        if not target:
            return None

        # target_ref 相对于 slide 文件（ppt/slides/）
        part_path = Path("ppt/slides") / target
        return normpath(str(part_path)).replace("\\", "/")
    except Exception:
        return None


def parse_shape(shape, slide=None, prs=None) -> Shape:
    """解析单个形状，按类型分发到对应解析器。

    Args:
        shape: python-pptx Shape 对象
        slide: 所属幻灯片（用于 placeholder 继承解析）
        prs: 所属 Presentation（用于主题色固化）
    """
    # 基础属性
    model = Shape(
        shape_id=str(getattr(shape, "shape_id", "") or ""),
        name=getattr(shape, "name", None),
    )

    # 通用属性
    props = _parse_common_props(shape, prs)
    for k, v in props.items():
        setattr(model, k, v)

    # 对 placeholder 形状，预先解析继承属性（对齐/字号/颜色/字体名），传入 prs 以固化主题色
    inherited_props = None
    if _is_placeholder(shape) and slide is not None:
        from ppt_transfor.utils.inheritance import resolve_placeholder_props
        inherited_props = resolve_placeholder_props(shape, slide, prs)
        # placeholder 类型决定默认对齐：标题类默认居中，正文类默认左对齐
        if inherited_props.get("alignment") is None:
            inherited_props["alignment"] = _placeholder_default_alignment(shape)

    # 类型判断与分发
    # 优先用 has_table / has_text_frame 等方法判断，更可靠
    if shape.has_table:
        model.shape_type = "table"
        from ppt_transfor.parser.table import parse_table
        model.table = parse_table(shape, prs)
        return model

    # shape_type 枚举判断
    try:
        st = shape.shape_type
        st_name = st.name if hasattr(st, "name") else str(st)
    except Exception:
        st = None
        st_name = "UNKNOWN"

    if st == MSO_SHAPE_TYPE.GROUP:
        model.shape_type = "group"
        from ppt_transfor.parser.group import parse_group
        group_fields = parse_group(shape, slide, prs)
        model.children = group_fields["children"]
        if group_fields.get("child_offset") is not None:
            model.child_offset = group_fields["child_offset"]
        if group_fields.get("child_extent") is not None:
            model.child_extent = group_fields["child_extent"]
        return model

    if st == MSO_SHAPE_TYPE.PICTURE:
        model.shape_type = "picture"
        from ppt_transfor.parser.image import parse_picture
        pic_fields = parse_picture(shape)
        for k, v in pic_fields.items():
            setattr(model, k, v)
        # 图片也可能有文本（虽然少见）
        if shape.has_text_frame:
            from ppt_transfor.parser.text import parse_text_frame
            model.text = parse_text_frame(shape.text_frame, prs, inherited_props)
        return model

    if st == MSO_SHAPE_TYPE.CHART:
        model.shape_type = "chart"
        try:
            model.chart_xml = etree.tostring(shape._element, encoding="unicode")
        except Exception:
            pass
        try:
            model.chart_part = _parse_chart_part(shape, slide)
        except Exception:
            pass
        return model

    if st in (MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.FREEFORM) and not shape.has_text_frame:
        # 连接线（无文本框的线）
        model.shape_type = "connector"
        from ppt_transfor.parser.connector import parse_connector
        conn_fields = parse_connector(shape)
        for k, v in conn_fields.items():
            setattr(model, k, v)
        return model

    if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
        model.shape_type = "auto_shape"
        from ppt_transfor.parser.autoshape import parse_autoshape
        auto_fields = parse_autoshape(shape)
        for k, v in auto_fields.items():
            setattr(model, k, v)
        # 自选图形通常有文本
        if shape.has_text_frame:
            from ppt_transfor.parser.text import parse_text_frame
            model.text = parse_text_frame(shape.text_frame, prs, inherited_props)
        return model

    if st == MSO_SHAPE_TYPE.TEXT_BOX:
        model.shape_type = "text_box"
        if shape.has_text_frame:
            from ppt_transfor.parser.text import parse_text_frame
            model.text = parse_text_frame(shape.text_frame, prs, inherited_props)
        return model

    if st == MSO_SHAPE_TYPE.PLACEHOLDER:
        model.shape_type = "placeholder"
        if shape.has_text_frame:
            from ppt_transfor.parser.text import parse_text_frame
            model.text = parse_text_frame(shape.text_frame, prs, inherited_props)
        return model

    # 兜底：未知类型，尝试提取文本
    model.shape_type = st_name.lower() if st_name else "unknown"
    if shape.has_text_frame:
        from ppt_transfor.parser.text import parse_text_frame
        model.text = parse_text_frame(shape.text_frame, prs, inherited_props)

    return model
