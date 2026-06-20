"""PPT 解析入口：Presentation → Presentation 模型。

打开 pptx，读取尺寸，遍历 slides。
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation as PptxPresentation

from ppt_transfor.models.schema import Presentation
from ppt_transfor.parser.slide import parse_slide


def parse_presentation(path: str | Path) -> Presentation:
    """解析 PPT 文件为 Presentation 模型

    Args:
        path: pptx 文件路径

    Returns:
        Presentation 模型
    """
    path = Path(path)
    prs = PptxPresentation(str(path))

    model = Presentation(
        source_file=path.name,
        slide_width=int(prs.slide_width),
        slide_height=int(prs.slide_height),
    )

    for idx, slide in enumerate(prs.slides):
        model.slides.append(parse_slide(slide, idx, prs))

    return model
