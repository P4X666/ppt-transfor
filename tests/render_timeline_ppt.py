"""使用 PowerPoint COM 将 timeline PPT 渲染为 PNG。"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import win32com.client
from win32com.client import constants


def render_pptx_to_png(pptx_path: Path, output_dir: Path, width: int = 1920, height: int = 1080) -> list[Path]:
    """使用 PowerPoint 将 PPTX 每页渲染为 PNG。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = False
    app.DisplayAlerts = False

    try:
        presentation = app.Presentations.Open(str(pptx_path.resolve()), WithWindow=False)
        try:
            images = []
            for i in range(1, presentation.Slides.Count + 1):
                slide = presentation.Slides.Item(i)
                out_path = output_dir / f"slide_{i:03d}.png"
                slide.Export(str(out_path.resolve()), "PNG", width, height)
                images.append(out_path)
            return images
        finally:
            presentation.Close()
    finally:
        app.Quit()


def main():
    out_dir = Path("out/visual_powerpoint/timeline_converted")
    images = render_pptx_to_png(Path("out/pptx/content_page__component_timeline.pptx"), out_dir)
    print(f"Rendered {len(images)} images to {out_dir}")
    for img in images:
        print(f"  {img}")


if __name__ == "__main__":
    main()
