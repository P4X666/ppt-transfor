"""幻灯片渲染器：Slide 模型 → Slide。

添加空白布局页，设置背景，遍历渲染 shapes。
"""

from __future__ import annotations

import logging

from ppt_transfor.models.schema import Background, Color, Slide
from ppt_transfor.renderer.shape import render_shape
from ppt_transfor.utils.color import apply_color

logger = logging.getLogger(__name__)


def _apply_background(slide, bg: Background | None) -> None:
    """设置幻灯片背景。

    保留 try-except 防止个别 slide 渲染失败中断整体流程，
    但增加日志警告便于定位问题。
    """
    if bg is None or bg.type == "none":
        return

    # solid 背景无颜色时不设置，避免默认黑色背景
    if bg.type == "solid" and bg.color is None:
        return

    try:
        fill = slide.background.fill
        if bg.type == "solid":
            fill.solid()
            if bg.color is not None:
                apply_color(fill.fore_color, bg.color)
    except Exception as e:
        logger.warning("背景渲染失败: %s, bg=%s", e, bg)


def _slide_bg_color(slide_model: Slide) -> Color | None:
    """从 Slide 模型的背景中提取 solid RGB 颜色，供 BACKGROUND 填充兜底使用。"""
    if slide_model.background is None:
        return None
    if slide_model.background.type != "solid":
        return None
    return slide_model.background.color


def render_slide(prs, slide_model: Slide, base_dir=None):
    """渲染单页幻灯片到 Presentation。

    Args:
        prs: python-pptx Presentation
        slide_model: Slide 模型
        base_dir: image_path 的基准目录（如 out/），传给图片渲染器

    Returns:
        python-pptx Slide 对象
    """
    # 使用空白布局（layouts[6] 通常是 Blank）
    # 兼容不同模板：优先按名称查找 Blank，找不到则用最后一个
    blank_layout = None
    for layout in prs.slide_layouts:
        if layout.name == "Blank":
            blank_layout = layout
            break
    if blank_layout is None:
        # 兜底：用最后一个布局（通常是 Blank）
        blank_layout = prs.slide_layouts[-1]

    slide = prs.slides.add_slide(blank_layout)

    # 背景
    _apply_background(slide, slide_model.background)

    # 提取 solid 背景色，供 BACKGROUND 填充形状兜底
    slide_bg_color = _slide_bg_color(slide_model)

    # 遍历渲染形状
    for shape_model in slide_model.shapes:
        render_shape(slide, shape_model, slide_bg_color, base_dir)

    return slide
