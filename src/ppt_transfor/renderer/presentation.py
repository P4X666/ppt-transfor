"""PPT 渲染入口：Presentation 模型 → PPT 文件。

创建 Presentation，设置尺寸，遍历 slides，保存。
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation as PptxPresentation
from pptx.util import Emu

from ppt_transfor.models.schema import Presentation
from ppt_transfor.renderer.slide import render_slide


def render_presentation(model: Presentation, output_path: str | Path) -> str:
    """渲染 Presentation 模型为 PPT 文件。

    Args:
        model: Presentation 模型
        output_path: 输出 pptx 文件路径

    Returns:
        实际保存的文件路径字符串
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    prs = PptxPresentation()

    # 设置幻灯片尺寸
    prs.slide_width = Emu(model.slide_width)
    prs.slide_height = Emu(model.slide_height)

    # 遍历渲染幻灯片
    for slide_model in model.slides:
        render_slide(prs, slide_model)

    prs.save(str(output_path))
    return str(output_path)
