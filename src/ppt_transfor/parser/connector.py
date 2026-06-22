"""连接线解析器：Connector → Connector 模型字段。

记录起点终点坐标与连接线样式。
"""

from __future__ import annotations


def parse_connector(shape) -> dict:
    """解析连接线，返回类型特有字段 dict。

    返回字段：
        begin_x/begin_y/end_x/end_y: 起终点坐标（EMU）

    对于 auto_shape 转来的连接线（如 prst="line"），shape.begin_x 等属性
    不可用，此时从 left/top/width/width/height 计算端点坐标，确保与
    render_connector 的兜底逻辑一致，避免往返差异。
    """
    fields = {}

    # 起终点坐标：优先用连接器原生属性，不可用时从几何边界计算
    try:
        begin_x = int(shape.begin_x) if shape.begin_x is not None else None
    except Exception:
        begin_x = None
    try:
        begin_y = int(shape.begin_y) if shape.begin_y is not None else None
    except Exception:
        begin_y = None
    try:
        end_x = int(shape.end_x) if shape.end_x is not None else None
    except Exception:
        end_x = None
    try:
        end_y = int(shape.end_y) if shape.end_y is not None else None
    except Exception:
        end_y = None

    # 兜底：从 left/top/width/height 计算端点（与 render_connector 逻辑一致）
    left = int(shape.left) if shape.left is not None else 0
    top = int(shape.top) if shape.top is not None else 0
    width = int(shape.width) if shape.width is not None else 0
    height = int(shape.height) if shape.height is not None else 0

    if begin_x is None:
        begin_x = left
    if begin_y is None:
        begin_y = top
    if end_x is None:
        end_x = left + width
    if end_y is None:
        end_y = top + height

    fields["begin_x"] = begin_x
    fields["begin_y"] = begin_y
    fields["end_x"] = end_x
    fields["end_y"] = end_y

    return fields
