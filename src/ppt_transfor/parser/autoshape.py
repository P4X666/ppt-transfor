"""自选图形解析器：AutoShape → AutoShape 模型字段。

记录 auto_shape_type 与 adjustments（调整手柄值，如圆角矩形的圆角比例）。
"""

from __future__ import annotations


def parse_autoshape(shape) -> dict:
    """解析自选图形，返回类型特有字段 dict。

    返回字段：
        auto_shape_type: 形状类型名（如 "ROUNDED_RECTANGLE"）
        adjustments: 调整手柄值列表
    """
    fields = {}

    # 形状类型
    try:
        ast = shape.auto_shape_type
        fields["auto_shape_type"] = ast.name if hasattr(ast, "name") else str(ast)
    except Exception:
        fields["auto_shape_type"] = None

    # 调整手柄值（如圆角矩形的圆角比例）
    adjustments = []
    try:
        for adj in shape.adjustments:
            # adj 可能是 float 或 None
            adjustments.append(float(adj) if adj is not None else None)
    except Exception:
        pass
    # 过滤 None 并去空
    adjustments = [a for a in adjustments if a is not None]
    if adjustments:
        fields["adjustments"] = adjustments

    return fields
