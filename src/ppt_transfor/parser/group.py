"""组合形状解析器：Group → Group 模型字段。

递归调用 shape.py 的 parse_shape 解析子形状。
同时捕获 group 的子坐标系 chOff/chExt（python-pptx 未暴露，需读 XML）。
"""

from __future__ import annotations

from lxml import etree

# DrawingML 命名空间
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _parse_group_xfrm(shape) -> tuple[tuple[int, int] | None, tuple[int, int] | None]:
    """解析 group shape 的 chOff/chExt。

    group 的 <p:grpSpPr>/<a:xfrm> 包含：
    - off/ext：组合在幻灯片上的位置（已由通用属性解析捕获）
    - chOff/chExt：子形状的坐标空间（本次新增）

    Args:
        shape: python-pptx GroupShape 对象

    Returns:
        (child_offset, child_extent)：
        - child_offset: (x, y) EMU 或 None
        - child_extent: (cx, cy) EMU 或 None
    """
    try:
        elem = shape._element
    except Exception:
        return None, None

    # 查找 <p:grpSpPr>/<a:xfrm>
    grp_sp_pr = elem.find(f"{{{NS_P}}}grpSpPr")
    if grp_sp_pr is None:
        return None, None

    xfrm = grp_sp_pr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return None, None

    child_offset = None
    child_extent = None

    # chOff：子坐标偏移
    ch_off = xfrm.find(f"{{{NS_A}}}chOff")
    if ch_off is not None:
        try:
            x = int(ch_off.get("x", "0"))
            y = int(ch_off.get("y", "0"))
            child_offset = (x, y)
        except (ValueError, TypeError):
            pass

    # chExt：子坐标范围
    ch_ext = xfrm.find(f"{{{NS_A}}}chExt")
    if ch_ext is not None:
        try:
            cx = int(ch_ext.get("cx", "0"))
            cy = int(ch_ext.get("cy", "0"))
            child_extent = (cx, cy)
        except (ValueError, TypeError):
            pass

    return child_offset, child_extent


def parse_group(shape, slide=None, prs=None, media_dir=None) -> dict:
    """解析组合形状，返回类型特有字段 dict。

    Args:
        shape: python-pptx GroupShape 对象
        slide: 所属幻灯片（传递给子形状解析）
        prs: 所属 Presentation（传递给子形状解析）
        media_dir: 图片输出目录（如 out/media），传递给子形状解析

    返回字段：
        children: 子形状列表
        child_offset: (x, y) EMU 或 None
        child_extent: (cx, cy) EMU 或 None
    """
    # 延迟导入避免循环依赖
    from ppt_transfor.parser.shape import parse_shape

    fields = {"children": [], "child_offset": None, "child_extent": None}

    # 解析子坐标系
    child_offset, child_extent = _parse_group_xfrm(shape)
    fields["child_offset"] = child_offset
    fields["child_extent"] = child_extent

    # 递归解析子形状
    for child in shape.shapes:
        fields["children"].append(parse_shape(child, slide, prs, media_dir))

    return fields
