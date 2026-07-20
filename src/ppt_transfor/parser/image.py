"""图片解析器：从 shape 提取图片，写入 out/media 文件，返回相对路径。

解析阶段把图片二进制提取为独立文件，JSON 中只存相对路径引用，
避免 base64 内嵌导致 JSON 膨胀。用内容 MD5 作文件名天然去重。

无 media_dir 时（如单元测试）降级为 base64 内嵌，保持向后兼容。
"""

from __future__ import annotations

import base64
import hashlib
from pathlib import Path

from ppt_transfor.models.schema import Crop


def parse_picture(shape, media_dir: Path | None = None) -> dict:
    """解析图片形状，写入文件返回相对路径。

    Args:
        shape: python-pptx Picture shape
        media_dir: 图片输出目录（如 out/media），None 时降级为 base64 内嵌

    Returns:
        字段 dict，含 image_path 或 data_base64、image_format、crop
    """
    fields = {}

    # 图片二进制数据
    try:
        image = shape.image
        blob = image.blob
        ext = image.ext or "png"

        if media_dir is not None:
            # 用内容 MD5 作文件名，天然去重（与 slide.py 去重逻辑一致）
            md5 = hashlib.md5(blob).hexdigest()
            filename = f"{md5}.{ext}"
            file_path = media_dir / filename

            # 幂等：已存在则跳过写入，避免重复 IO
            if not file_path.exists():
                media_dir.mkdir(parents=True, exist_ok=True)
                file_path.write_bytes(blob)

            # 存相对路径（相对 out/ 目录，JSON 与 media 同父目录便于整体迁移）
            fields["image_path"] = f"media/{filename}"
            fields["image_format"] = ext
        else:
            # 无 media_dir 时降级为 base64 内嵌（向后兼容/测试场景）
            fields["data_base64"] = base64.b64encode(blob).decode("ascii")
            fields["image_format"] = ext
    except Exception:
        # 无法读取图片数据（可能是链接图片等），静默跳过
        fields["data_base64"] = None
        fields["image_path"] = None
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
