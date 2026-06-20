"""形状渲染器：通用属性回写 + 按类型分发。

通用属性：位置/旋转/填充/边框/阴影
类型分发：table/group/picture/connector/autoshape/text_box/placeholder
"""

from __future__ import annotations

from pptx.enum.dml import MSO_FILL
from pptx.util import Emu

from ppt_transfor.models.schema import Fill, Line, Shape
from ppt_transfor.utils.color import apply_color


def _apply_fill(fill, fill_model: Fill | None) -> None:
    """回写填充"""
    if fill_model is None:
        return

    if fill_model.type == "none":
        try:
            fill.background()
        except Exception:
            pass
        return

    if fill_model.type == "solid":
        try:
            fill.solid()
            if fill_model.color is not None:
                apply_color(fill.fore_color, fill_model.color)
        except Exception:
            pass
        return

    if fill_model.type == "gradient":
        # 基本渐变回写：设置 gradient 类型，stops 保留默认（完整 stops 回写待扩展）
        try:
            fill.gradient()
        except Exception:
            pass
        return

    # pattern/picture 等高级填充暂不支持回写
    # 后续可通过 xml_helper 直接操作 XML 扩展


def _apply_line(line, line_model: Line | None) -> None:
    """回写边框"""
    if line_model is None:
        return

    if line_model.width is not None:
        try:
            line.width = Emu(line_model.width)
        except Exception:
            pass
    if line_model.color is not None:
        try:
            apply_color(line.color, line_model.color)
        except Exception:
            pass
    if line_model.dash is not None:
        try:
            from pptx.enum.dml import MSO_LINE_DASH_STYLE
            line.dash_style = MSO_LINE_DASH_STYLE[line_model.dash]
        except (KeyError, ValueError, Exception):
            pass


def _apply_common_props(shape, model: Shape) -> None:
    """回写通用属性（位置/旋转/填充/边框/阴影）"""
    # 位置与尺寸
    if model.left is not None:
        try:
            shape.left = Emu(model.left)
        except Exception:
            pass
    if model.top is not None:
        try:
            shape.top = Emu(model.top)
        except Exception:
            pass
    if model.width is not None:
        try:
            shape.width = Emu(model.width)
        except Exception:
            pass
    if model.height is not None:
        try:
            shape.height = Emu(model.height)
        except Exception:
            pass

    # 旋转
    if model.rotation:
        try:
            shape.rotation = model.rotation
        except Exception:
            pass

    # 填充
    try:
        _apply_fill(shape.fill, model.fill)
    except Exception:
        pass

    # 边框
    try:
        _apply_line(shape.line, model.line)
    except Exception:
        pass

    # 阴影：若模型未指定，显式关闭继承，避免 add_shape 默认继承主题阴影
    if model.shadow is not None:
        try:
            shape.shadow.inherit = model.shadow.enabled
        except Exception:
            pass
    else:
        try:
            shape.shadow.inherit = False
        except Exception:
            pass


def render_shape(container, model: Shape):
    """渲染单个形状到容器（slide 或 group），按类型分发。

    Args:
        container: python-pptx Slide 或 GroupShape
        model: Shape 模型

    Returns:
        渲染后的 python-pptx 形状对象
    """
    shape = None

    if model.shape_type == "table":
        from ppt_transfor.renderer.table import render_table
        shape = render_table(container, model)
    elif model.shape_type == "group":
        from ppt_transfor.renderer.group import render_group
        shape = render_group(container, model)
    elif model.shape_type == "picture":
        from ppt_transfor.renderer.image import render_picture
        shape = render_picture(container, model)
    elif model.shape_type == "connector":
        from ppt_transfor.renderer.connector import render_connector
        shape = render_connector(container, model)
    elif model.shape_type == "auto_shape":
        from ppt_transfor.renderer.autoshape import render_autoshape
        shape = render_autoshape(container, model)
        # auto_shape_type 为 None 或无法识别时，降级为文本框保持位置和尺寸
        if shape is None:
            from ppt_transfor.renderer.text import render_text_frame
            shape = container.shapes.add_textbox(
                Emu(model.left) if model.left is not None else None,
                Emu(model.top) if model.top is not None else None,
                Emu(model.width) if model.width is not None else None,
                Emu(model.height) if model.height is not None else None,
            )
            if model.text is not None:
                render_text_frame(shape.text_frame, model.text)
    elif model.shape_type in ("text_box", "placeholder"):
        # 文本框：add_textbox 后回写文本
        from ppt_transfor.renderer.text import render_text_frame
        shape = container.shapes.add_textbox(
            Emu(model.left) if model.left is not None else None,
            Emu(model.top) if model.top is not None else None,
            Emu(model.width) if model.width is not None else None,
            Emu(model.height) if model.height is not None else None,
        )
        if model.text is not None:
            render_text_frame(shape.text_frame, model.text)
    else:
        # 兜底：未知类型（如 chart），渲染占位文本框保持位置和尺寸，
        # 确保形状数量与顺序不变（chart 内容本身不支持往返）
        from ppt_transfor.renderer.text import render_text_frame
        shape = container.shapes.add_textbox(
            Emu(model.left) if model.left is not None else None,
            Emu(model.top) if model.top is not None else None,
            Emu(model.width) if model.width is not None else None,
            Emu(model.height) if model.height is not None else None,
        )
        if model.text is not None:
            render_text_frame(shape.text_frame, model.text)

    # 回写通用属性（图片/表格的位置已在 add_xxx 时设置，这里再次确保）
    if shape is not None:
        _apply_common_props(shape, model)

        # 自选图形/文本框的文本回写（auto_shape 在 add_shape 时已创建空文本框）
        if model.shape_type == "auto_shape" and model.text is not None:
            from ppt_transfor.renderer.text import render_text_frame
            try:
                render_text_frame(shape.text_frame, model.text)
            except Exception:
                pass

    return shape
