"""检查转换后的 timeline PPT 的幻灯片背景。"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from pptx import Presentation as PptxPresentation


def main():
    pptx_path = Path("out/pptx/content_page__component_timeline.pptx")
    print("=== CONVERTED PPTX BACKGROUNDS ===")
    prs = PptxPresentation(str(pptx_path))
    for si, slide in enumerate(prs.slides):
        print(f"Slide {si}:")
        bg = slide.background
        fill = bg.fill
        print(f"  bg.fill.type={fill.type}")
        try:
            print(f"  bg.fill.fore_color.rgb={fill.fore_color.rgb}")
        except Exception as e:
            print(f"  bg.fill.fore_color error={e}")
        try:
            print(f"  bg.fill.back_color.rgb={fill.back_color.rgb}")
        except Exception as e:
            print(f"  bg.fill.back_color error={e}")


if __name__ == "__main__":
    main()
