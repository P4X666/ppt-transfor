"""组合形状渲染器：Group 模型 → Group 形状。

递归渲染子形状后组合，并回写子坐标系 chOff/chExt。
"""

from __future__ import annotations

from lxml import etree
from pptx.util import Emu

from ppt_transfor.models.schema import Shape

# DrawingML 命名空间
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _apply_child_coords(group, model: Shape) -> None:
    """回写 group 的子坐标系 chOff/chExt。

    python-pptx 未暴露 chOff/chExt API，需直接操作 XML。
    group 的 <p:grpSpPr>/<a:xfrm> 下：
    - <a:chOff x="..." y="..."/>
    - <a:chExt cx="..." cy="..."/>

    Args:
        group: python-pptx GroupShape 对象
        model: Shape 模型（含 child_offset/child_extent）
    """
    if model.child_offset is None and model.child_extent is None:
        return

    try:
        elem = group._element
    except Exception:
        return

    # 查找或创建 <p:grpSpPr>
    grp_sp_pr = elem.find(f"{{{NS_P}}}grpSpPr")
    if grp_sp_pr is None:
        grp_sp_pr = etree.SubElement(elem, f"{{{NS_P}}}grpSpPr")

    # 查找或创建 <a:xfrm>
    xfrm = grp_sp_pr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        xfrm = etree.SubElement(grp_sp_pr, f"{{{NS_A}}}xfrm")

    # 回写 chOff
    if model.child_offset is not None:
        ch_off = xfrm.find(f"{{{NS_A}}}chOff")
        if ch_off is None:
            ch_off = etree.SubElement(xfrm, f"{{{NS_A}}}chOff")
        ch_off.set("x", str(model.child_offset[0]))
        ch_off.set("y", str(model.child_offset[1]))

    # 回写 chExt
    if model.child_extent is not None:
        ch_ext = xfrm.find(f"{{{NS_A}}}chExt")
        if ch_ext is None:
            ch_ext = etree.SubElement(xfrm, f"{{{NS_A}}}chExt")
        ch_ext.set("cx", str(model.child_extent[0]))
        ch_ext.set("cy", str(model.child_extent[1]))


def render_group(slide, model: Shape, slide_bg_color=None):
    """渲染组合形状，返回 python-pptx GroupShape 对象。

    Args:
        slide: python-pptx Slide
        model: Shape 模型（shape_type == "group"）
        slide_bg_color: 当前幻灯片背景色，传递给子形状

    Returns:
        GroupShape 对象
    """
    # 延迟导入避免循环依赖
    from ppt_transfor.renderer.shape import render_shape

    # 先用 add_group_shape 创建空组合（python-pptx 1.0+ 支持）
    # 若版本不支持，则降级为依次添加子形状（不组合）
    try:
        group = slide.shapes.add_group_shape()
    except (AttributeError, TypeError):
        # 不支持 add_group_shape，降级渲染第一个子形状作为占位
        for child in model.children:
            render_shape(slide, child)
        return None

    # 设置组合的位置和尺寸
    if model.left is not None:
        group.left = Emu(model.left)
    if model.top is not None:
        group.top = Emu(model.top)
    if model.width is not None:
        group.width = Emu(model.width)
    if model.height is not None:
        group.height = Emu(model.height)

    # 递归渲染子形状到组合内
    for child in model.children:
        render_shape(group, child, slide_bg_color)

    # 回写子坐标系 chOff/chExt（关键：在添加子形状之后回写，
    # 避免 python-pptx 添加子形状时自动重算 chOff/chExt 覆盖原始值）
    _apply_child_coords(group, model)

    return group
