"""EMU 单位换算工具。

python-pptx 内部统一用 EMU（914400 EMU = 1 英寸 = 2.54cm = 72pt）。
JSON 中所有几何尺寸存 EMU 整数，避免浮点误差。
"""

# EMU 换算常量
EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700


def emu_to_pt(emu: int | None) -> float | None:
    """EMU → pt"""
    return None if emu is None else emu / EMU_PER_PT


def pt_to_emu(pt: float) -> int:
    """pt → EMU"""
    return int(round(pt * EMU_PER_PT))


def emu_to_cm(emu: int | None) -> float | None:
    """EMU → cm"""
    return None if emu is None else emu / EMU_PER_CM


def cm_to_emu(cm: float) -> int:
    """cm → EMU"""
    return int(round(cm * EMU_PER_CM))


def emu_to_inch(emu: int | None) -> float | None:
    """EMU → inch"""
    return None if emu is None else emu / EMU_PER_INCH


def inch_to_emu(inch: float) -> int:
    """inch → EMU"""
    return int(round(inch * EMU_PER_INCH))
