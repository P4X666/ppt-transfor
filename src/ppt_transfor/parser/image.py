"""图片解析器：Picture 形状 → Picture 模型。

读取图片二进制数据转 base64，记录格式与裁剪信息。
"""

from __future__ import annotations

import base64

from ppt_transfor.models.schema import Crop


def parse_picture(shape) -> dict:
    """解析图片形状，返回类型特有字段 dict。

    返回字段：
        data_base64: 图片二进制 base64 编码
        image_format: 图片格式（png/jpeg/gif 等）
        crop: 裁剪信息
    """
    fields = {}

    # 图片二进制数据
    try:
        image = shape.image
        blob = image.blob
        fields["data_base64"] = base64.b64encode(blob).decode("ascii")
        # 图片格式：image.ext 不带点（如 "png"）
        fields["image_format"] = image.ext
    except Exception:
        # 无法读取图片数据（可能是链接图片等）
        fields["data_base64"] = None
        fields["image_format"] = None

    # 裁剪信息
    try:
        crop = Crop(
            left=float(shape.crop_left or 0.0),
            top=float(shape.crop_top or 0.0),
            right=float(shape.crop_right or 0.0),
            bottom=float(shape.crop_bottom or 0.0),
        )
        # 全 0 时不记录，减少噪声
        if crop.left or crop.top or crop.right or crop.bottom:
            fields["crop"] = crop
    except Exception:
        pass

    return fields
