"""形状解析器：通用属性提取 + 按类型分发。

通用属性：shape_id/name/shape_type/位置/旋转/填充/边框/阴影
类型分发：table/group/picture/connector/autoshape/text_box/placeholder
对 placeholder 形状，解析继承自 layout/master 的对齐、字号、颜色。
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE_TYPE

from ppt_transfor.models.schema import Fill, Line, Shadow, Shape
from ppt_transfor.utils.color import parse_color


def _parse_fill(fill, prs=None) -> Fill | None:
    """解析填充。

    None（继承）与 BACKGROUND（显式无填充）统一返回 None，
    避免往返时因 add_shape 默认填充行为差异产生噪声。
    """
    try:
        fill_type = fill.type
    except Exception:
        return None

    if fill_type is None:
        return None  # 继承样式，不记录

    # fill_type 是 MSO_FILL 枚举
    type_name = fill_type.name if hasattr(fill_type, "name") else str(fill_type)

    # BACKGROUND（显式无填充）与继承统一，往返一致
    if type_name == "BACKGROUND":
        return None

    model = Fill()
    if type_name == "SOLID":
        model.type = "solid"
        try:
            color = parse_color(fill.fore_color, prs)
            if color is not None:
                model.color = color
        except Exception:
            pass
    elif type_name == "GRADIENT":
        model.type = "gradient"
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
        fill = _parse_fill(shape.fill, prs)
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
