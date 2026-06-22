"""颜色解析工具：将 python-pptx 颜色对象统一转为 Color 模型。

python-pptx 颜色来源多样：
- RGBColor（直接 RGB，如 RGBColor(0xFF, 0x00, 0x00)）
- 主题色（MSO_THEME_COLOR.ACCENT1 等）
- scheme 色

主题色固化为 RGB：解析时通过 theme part 的 clrScheme 把主题色解析为具体 RGB，
牺牲主题色语义换 100% 视觉保真（避免新 PPT 主题色索引映射错位）。
"""

from __future__ import annotations

from typing import Optional

from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

from ppt_transfor.models.schema import Color


def parse_color(color_format, prs=None) -> Optional[Color]:
    """解析 python-pptx 的 ColorFormat 对象为 Color 模型。

    主题色（SCHEME 类型）固化为 RGB：通过 prs 的 theme part 解析为具体 RGB 值。

    Args:
        color_format: font.color 或 fill.fore_color 等 ColorFormat 对象
        prs: 所属 Presentation 对象（用于主题色固化，可选）

    Returns:
        Color 模型，若颜色未设置（type=None）则返回 None
    """
    if color_format is None:
        return None

    try:
        color_type = color_format.type
    except Exception:
        return None

    if color_type is None:
        return None

    # RGB 直接颜色
    if str(color_type) == "RGB" or color_type == 1:
        try:
            rgb: RGBColor = color_format.rgb
            return Color(type="rgb", value=str(rgb))
        except Exception:
            return None

    # 主题色：固化为 RGB
    if str(color_type) == "SCHEME" or color_type == 2:
        # 优先尝试固化为主题色对应的 RGB
        if prs is not None:
            from ppt_transfor.utils.inheritance import resolve_theme_color
            rgb_color = resolve_theme_color(color_format, prs)
            if rgb_color is not None:
                return rgb_color
        # 固化失败时降级保留主题色名（兜底），避免颜色完全丢失导致默认黑色
        try:
            theme_color = color_format.theme_color
            return Color(
                type="theme",
                value=theme_color.name if hasattr(theme_color, "name") else str(theme_color),
            )
        except Exception:
            return None

    # 其他类型尝试取 rgb
    try:
        rgb = color_format.rgb
        return Color(type="rgb", value=str(rgb))
    except Exception:
        return None


def apply_color(color_format, color: Optional[Color]) -> None:
    """将 Color 模型回写到 python-pptx 的 ColorFormat 对象。

    Args:
        color_format: font.color 或 fill.fore_color 等 ColorFormat 对象
        color: Color 模型，None 则跳过
    """
    if color is None:
        return

    if color.type == "rgb":
        # 解析 "FF0000" 为 RGBColor
        hex_str = color.value.lstrip("#")
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
        color_format.rgb = RGBColor(r, g, b)
    elif color.type == "theme":
        # 按名称查找 MSO_THEME_COLOR 枚举
        try:
            theme_color = MSO_THEME_COLOR.from_xml(color.value)
        except (ValueError, KeyError):
            # 尝试按成员名匹配
            for member in MSO_THEME_COLOR:
                if member.name == color.value:
                    theme_color = member
                    break
            else:
                return
        color_format.theme_color = theme_color
    # scheme 类型暂不支持回写（需更复杂的主题映射）
