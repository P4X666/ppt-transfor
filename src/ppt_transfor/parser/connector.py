"""连接线解析器：Connector → Connector 模型字段。

记录起点终点坐标与连接线样式。
"""

from __future__ import annotations


def parse_connector(shape) -> dict:
    """解析连接线，返回类型特有字段 dict。

    返回字段：
        begin_x/begin_y/end_x/end_y: 起终点坐标（EMU）
    """
    fields = {}

    # 起终点坐标
    try:
        fields["begin_x"] = int(shape.begin_x) if shape.begin_x is not None else None
    except Exception:
        fields["begin_x"] = None
    try:
        fields["begin_y"] = int(shape.begin_y) if shape.begin_y is not None else None
    except Exception:
        fields["begin_y"] = None
    try:
        fields["end_x"] = int(shape.end_x) if shape.end_x is not None else None
    except Exception:
        fields["end_x"] = None
    try:
        fields["end_y"] = int(shape.end_y) if shape.end_y is not None else None
    except Exception:
        fields["end_y"] = None

    return fields
