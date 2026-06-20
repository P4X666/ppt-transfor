"""图片渲染器：Picture 模型 → Picture 形状。

base64 解码为 blob，add_picture 添加，回写裁剪。
"""

from __future__ import annotations

import base64
import io

from pptx.util import Emu

from ppt_transfor.models.schema import Shape


def render_picture(slide, model: Shape):
    """渲染图片形状，返回 python-pptx Picture 对象。

    Args:
        slide: python-pptx Slide
        model: Shape 模型（shape_type == "picture"）

    Returns:
        Picture 对象；若图片数据缺失则返回 None
    """
    if not model.data_base64:
        return None

    blob = base64.b64decode(model.data_base64)
    image_stream = io.BytesIO(blob)

    pic = slide.shapes.add_picture(
        image_stream,
        Emu(model.left) if model.left is not None else None,
        Emu(model.top) if model.top is not None else None,
        Emu(model.width) if model.width is not None else None,
        Emu(model.height) if model.height is not None else None,
    )

    # 裁剪
    if model.crop is not None:
        try:
            pic.crop_left = model.crop.left
            pic.crop_top = model.crop.top
            pic.crop_right = model.crop.right
            pic.crop_bottom = model.crop.bottom
        except Exception:
            pass

    return pic
