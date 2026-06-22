"""检查 card PPT 的形状结构，看是否有组合。"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from ppt_transfor.parser.presentation import parse_presentation


def print_shape(s, indent=0):
    prefix = "  " * indent
    print(f"{prefix}Shape: {getattr(s, 'name', None)} type={type(s).__name__}")
    print(f"{prefix}  left={s.left} top={s.top} width={s.width} height={s.height}")
    print(f"{prefix}  fill={getattr(s, 'fill', None)}")
    if getattr(s, "text", None):
        for pi, para in enumerate(s.text.paragraphs):
            for ri, run in enumerate(para.runs):
                if run.text.strip():
                    print(f"{prefix}  text[{pi},{ri}]: '{run.text[:30]}' color={run.font.color}")
    for child in getattr(s, "children", []) or []:
        print_shape(child, indent + 1)


def main():
    model = parse_presentation("input/conent_page_component_card.pptx")
    for si, slide in enumerate(model.slides):
        print(f"\n=== Slide {si} ===")
        for shape in slide.shapes:
            print_shape(shape)


if __name__ == "__main__":
    main()
