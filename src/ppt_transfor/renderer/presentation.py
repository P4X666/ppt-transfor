"""PPT 渲染入口：Presentation 模型 → PPT 文件。

创建 Presentation，设置尺寸，遍历 slides，保存。
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation as PptxPresentation
from pptx.util import Emu

from ppt_transfor.models.schema import Presentation
from ppt_transfor.renderer.slide import render_slide


def render_presentation(
    model: Presentation,
    output_path: str | Path,
    source_pptx_path: str | Path | None = None,
) -> str:
    """渲染 Presentation 模型为 PPT 文件。

    Args:
        model: Presentation 模型
        output_path: 输出 pptx 文件路径
        source_pptx_path: 原始 PPTX 路径，若提供则在保存后尝试保留 chart 等 XML 级内容

    Returns:
        实际保存的文件路径字符串
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # image_path 的基准目录：输出 PPTX 的祖父目录（如 out/）
    # 解析时图片存到 out/media/abc.png，JSON 中存 "media/abc.png"
    # 渲染时 output_path=out/pptx/xxx.pptx → base_dir=out/
    # 这样 base_dir/image_path = out/media/abc.png 能正确解析
    base_dir = output_path.parent.parent

    prs = PptxPresentation()

    # 设置幻灯片尺寸
    prs.slide_width = Emu(model.slide_width)
    prs.slide_height = Emu(model.slide_height)

    # 遍历渲染幻灯片
    for slide_model in model.slides:
        render_slide(prs, slide_model, base_dir)

    prs.save(str(output_path))

    # 若存在原始文件，先执行主题保留（替换 theme1.xml 和 tableStyles.xml），
    # 再执行 chart 保留后处理
    if source_pptx_path is None and model.source_file:
        inferred = Path("input") / model.source_file
        if inferred.exists():
            source_pptx_path = inferred

    if source_pptx_path is not None:
        try:
            from ppt_transfor.renderer.theme_preserve import apply_theme_preservation

            apply_theme_preservation(output_path, source_pptx_path)
        except Exception:
            # 主题保留失败不应影响主流程
            pass

        try:
            from ppt_transfor.renderer.chart_preserve import apply_chart_preservation

            apply_chart_preservation(output_path, source_pptx_path, model)
        except Exception:
            # chart 保留失败不应影响主流程
            pass

    return str(output_path)
