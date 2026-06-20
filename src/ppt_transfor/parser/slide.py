"""幻灯片解析器：Slide → Slide 模型。

解析背景（含继承链）+ 遍历 shapes。
"""

from __future__ import annotations

from ppt_transfor.models.schema import Slide
from ppt_transfor.parser.shape import parse_shape
from ppt_transfor.utils.inheritance import resolve_background


def parse_slide(slide, index: int, prs=None) -> Slide:
    """解析单页幻灯片

    Args:
        slide: python-pptx Slide 对象
        index: 幻灯片索引
        prs: 所属 Presentation 对象（传递给 shape 解析用于继承解析）
    """
    model = Slide(index=index)

    # 布局名
    try:
        model.layout_name = slide.slide_layout.name
    except Exception:
        model.layout_name = "Blank"

    # 背景：沿继承链解析（slide → layout → master），传入 prs 以固化主题色
    bg = resolve_background(slide, prs)
    if bg is not None:
        model.background = bg

    # 遍历形状
    for shape in slide.shapes:
        model.shapes.append(parse_shape(shape, slide, prs))

    return model
