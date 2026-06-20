"""连接线渲染器：Connector 模型 → Connector 形状。

add_connector 后设置起终点坐标。
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_CONNECTOR
from pptx.util import Emu

from ppt_transfor.models.schema import Shape


def render_connector(slide, model: Shape):
    """渲染连接线，返回 python-pptx Connector 对象。

    Args:
        slide: python-pptx Slide
        model: Shape 模型（shape_type == "connector"）

    Returns:
        Connector 对象
    """
    # 默认用直线连接器
    begin_x = model.begin_x if model.begin_x is not None else (model.left or 0)
    begin_y = model.begin_y if model.begin_y is not None else (model.top or 0)
    end_x = model.end_x if model.end_x is not None else ((model.left or 0) + (model.width or 0))
    end_y = model.end_y if model.end_y is not None else ((model.top or 0) + (model.height or 0))

    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Emu(begin_x),
        Emu(begin_y),
        Emu(end_x),
        Emu(end_y),
    )

    return connector
