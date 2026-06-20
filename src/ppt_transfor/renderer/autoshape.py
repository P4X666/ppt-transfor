"""自选图形渲染器：AutoShape 模型 → AutoShape 形状。

add_shape(MSO_SHAPE.XXX) 后回写 adjustments。
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu

from ppt_transfor.models.schema import Shape


def render_autoshape(slide, model: Shape):
    """渲染自选图形，返回 python-pptx Shape 对象。

    Args:
        slide: python-pptx Slide
        model: Shape 模型（shape_type == "auto_shape"）

    Returns:
        Shape 对象；若类型无法识别则返回 None
    """
    if not model.auto_shape_type:
        return None

    # 按名称查找 MSO_SHAPE 枚举成员
    shape_enum = None
    try:
        shape_enum = MSO_SHAPE[model.auto_shape_type]
    except KeyError:
        # 尝试大小写不敏感匹配
        for member in MSO_SHAPE:
            if member.name == model.auto_shape_type:
                shape_enum = member
                break
    if shape_enum is None:
        return None

    shape = slide.shapes.add_shape(
        shape_enum,
        Emu(model.left) if model.left is not None else None,
        Emu(model.top) if model.top is not None else None,
        Emu(model.width) if model.width is not None else None,
        Emu(model.height) if model.height is not None else None,
    )

    # 回写 adjustments（调整手柄值）
    for idx, adj_value in enumerate(model.adjustments):
        try:
            if idx < len(shape.adjustments):
                shape.adjustments[idx] = adj_value
        except Exception:
            pass

    return shape
