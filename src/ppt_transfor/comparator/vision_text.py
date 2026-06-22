"""基于解析模型的文本可见性视觉检测。

不依赖 LibreOffice/PowerPoint 渲染，而是：
1. 解析 PPTX 为内部模型（已解析继承/主题色）
2. 计算每个文本 run 的有效字体颜色与有效背景颜色
3. 计算亮度对比度，标记低对比度（黑底黑字/白底白字）
4. 用 Pillow 生成可视化诊断图

适用于批量扫描所有 PPT，定位文本不可见问题。
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from ppt_transfor.models.schema import Background, Color, Fill, Presentation, Run, Shape, Slide


# 低对比度阈值：亮度差低于此值视为不可见
CONTRAST_THRESHOLD = 0.2


@dataclass
class VisibilityIssue:
    """文本可见性问题记录。"""

    file_name: str
    slide_index: int
    shape_name: Optional[str]
    text: str
    text_color: Optional[str]  # RRGGBB or None
    bg_color: Optional[str]  # RRGGBB or None
    text_lum: float
    bg_lum: float
    contrast: float
    reason: str


def _color_to_rgb(color: Color | None) -> str | None:
    """Color 模型转 RRGGBB，仅处理 rgb 类型。"""
    if color is None:
        return None
    if color.type == "rgb":
        val = color.value.lstrip("#")
        if len(val) == 6:
            return val
    return None


def _rgb_to_luminance(hex_rgb: str) -> float:
    """RRGGBB → 相对亮度（0-1）。"""
    hex_rgb = hex_rgb.lstrip("#")
    if len(hex_rgb) != 6:
        return 0.5
    r = int(hex_rgb[0:2], 16) / 255
    g = int(hex_rgb[2:4], 16) / 255
    b = int(hex_rgb[4:6], 16) / 255
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _effective_bg_rgb(slide: Slide, shape: Shape) -> str | None:
    """计算形状的有效背景 RGB。"""
    fill = getattr(shape, "fill", None)
    if fill is not None:
        if fill.type == "solid":
            rgb = _color_to_rgb(fill.color)
            if rgb:
                return rgb
        # background 类型跟随 slide bg
    bg = slide.background
    if bg is not None and bg.type == "solid":
        rgb = _color_to_rgb(bg.color)
        if rgb:
            return rgb
    return None


def _effective_text_rgb(run: Run) -> str | None:
    """计算 run 的有效字体颜色 RGB。"""
    return _color_to_rgb(run.font.color)


def _detect_shape_text_issues(
    file_name: str,
    slide: Slide,
    shape: Shape,
    threshold: float = CONTRAST_THRESHOLD,
) -> list[VisibilityIssue]:
    """检测单个 shape 内文本可见性问题。"""
    issues: list[VisibilityIssue] = []
    bg_rgb = _effective_bg_rgb(slide, shape)
    bg_lum = _rgb_to_luminance(bg_rgb) if bg_rgb else 0.5

    text = getattr(shape, "text", None)
    if text is None:
        return issues

    for para in text.paragraphs:
        for run in para.runs:
            run_text = run.text.strip()
            if not run_text:
                continue

            text_rgb = _effective_text_rgb(run)
            if text_rgb is None:
                # 未显式设置颜色，默认按黑色处理；若背景很深则不可见
                text_lum = _rgb_to_luminance("000000")
                contrast = abs(text_lum - bg_lum)
                if bg_lum < 0.3 and contrast < threshold:
                    issues.append(
                        VisibilityIssue(
                            file_name=file_name,
                            slide_index=slide.index,
                            shape_name=shape.name,
                            text=run_text,
                            text_color=None,
                            bg_color=bg_rgb,
                            text_lum=text_lum,
                            bg_lum=bg_lum,
                            contrast=contrast,
                            reason="默认黑字在深色背景上",
                        )
                    )
                continue

            text_lum = _rgb_to_luminance(text_rgb)
            contrast = abs(text_lum - bg_lum)
            if contrast < threshold:
                issues.append(
                    VisibilityIssue(
                        file_name=file_name,
                        slide_index=slide.index,
                        shape_name=shape.name,
                        text=run_text,
                        text_color=text_rgb,
                        bg_color=bg_rgb,
                        text_lum=text_lum,
                        bg_lum=bg_lum,
                        contrast=contrast,
                        reason=f"亮度差 {contrast:.2f} 低于阈值 {threshold}",
                    )
                )

    return issues


def _walk_shapes(shape: Shape) -> list[Shape]:
    """递归遍历 shape（含组合子形状）。"""
    shapes = [shape]
    for child in getattr(shape, "children", []) or []:
        shapes.extend(_walk_shapes(child))
    return shapes


def _detect_table_text_issues(
    file_name: str,
    slide: Slide,
    shape: Shape,
    threshold: float = CONTRAST_THRESHOLD,
) -> list[VisibilityIssue]:
    """检测表格单元格内文本可见性问题。"""
    issues: list[VisibilityIssue] = []
    table = getattr(shape, "table", None)
    if table is None:
        return issues

    for row in table.cells:
        for cell in row:
            cell_bg_rgb = None
            if cell.fill is not None and cell.fill.type == "solid":
                cell_bg_rgb = _color_to_rgb(cell.fill.color)
            if cell_bg_rgb is None:
                cell_bg_rgb = _effective_bg_rgb(slide, shape)
            bg_lum = _rgb_to_luminance(cell_bg_rgb) if cell_bg_rgb else 0.5

            for para in cell.text.paragraphs:
                for run in para.runs:
                    run_text = run.text.strip()
                    if not run_text:
                        continue
                    text_rgb = _effective_text_rgb(run)
                    if text_rgb is None:
                        text_rgb = "000000"
                    text_lum = _rgb_to_luminance(text_rgb)
                    contrast = abs(text_lum - bg_lum)
                    if contrast < threshold:
                        issues.append(
                            VisibilityIssue(
                                file_name=file_name,
                                slide_index=slide.index,
                                shape_name=f"{shape.name} 表格单元格",
                                text=run_text,
                                text_color=text_rgb,
                                bg_color=cell_bg_rgb,
                                text_lum=text_lum,
                                bg_lum=bg_lum,
                                contrast=contrast,
                                reason=f"表格内亮度差 {contrast:.2f} 低于阈值 {threshold}",
                            )
                        )
    return issues


def detect_presentation_visibility_issues(
    model: Presentation,
    threshold: float = CONTRAST_THRESHOLD,
) -> list[VisibilityIssue]:
    """扫描整个 Presentation 模型，返回所有低对比度文本问题。"""
    issues: list[VisibilityIssue] = []
    file_name = Path(model.source_file).name if model.source_file else "unknown"

    for slide in model.slides:
        for shape in slide.shapes:
            for s in _walk_shapes(shape):
                issues.extend(_detect_shape_text_issues(file_name, slide, s, threshold))
                issues.extend(_detect_table_text_issues(file_name, slide, s, threshold))

    return issues


def _hex_to_fill(hex_rgb: str | None) -> str:
    """RRGGBB → #RRGGBB。"""
    if hex_rgb is None:
        return "#808080"
    return f"#{hex_rgb}"


def _contrast_bar_color(text_lum: float, bg_lum: float) -> str:
    """根据对比度返回状态条颜色：绿/黄/红。"""
    diff = abs(text_lum - bg_lum)
    if diff >= 0.4:
        return "#22c55e"
    if diff >= 0.2:
        return "#eab308"
    return "#ef4444"


def generate_visibility_report_image(issues: list[VisibilityIssue], output_path: Path) -> None:
    """生成文本可见性问题汇总图（长图）。"""
    from PIL import Image, ImageDraw, ImageFont

    row_height = 70
    padding = 20
    width = 1200
    height = padding * 2 + max(len(issues), 1) * row_height + 60

    img = Image.new("RGB", (width, height), "#f8fafc")
    draw = ImageDraw.Draw(img)

    try:
        font = ImageFont.truetype("arial.ttf", 18)
        small_font = ImageFont.truetype("arial.ttf", 14)
        title_font = ImageFont.truetype("arial.ttf", 24)
    except Exception:
        font = ImageFont.load_default()
        small_font = font
        title_font = font

    draw.text((padding, padding), "文本可见性检测报告", fill="#0f172a", font=title_font)
    draw.text(
        (padding, padding + 35),
        f"共发现 {len(issues)} 处低对比度问题（阈值 0.2）",
        fill="#475569",
        font=small_font,
    )

    y = padding + 70
    for issue in issues:
        # 文本颜色方块
        draw.rectangle([padding, y + 10, padding + 40, y + 50], fill=_hex_to_fill(issue.text_color), outline="#334155")
        # 背景颜色方块
        draw.rectangle([padding + 50, y + 10, padding + 90, y + 50], fill=_hex_to_fill(issue.bg_color), outline="#334155")

        # 文本信息
        text_x = padding + 110
        draw.text((text_x, y + 5), f"{issue.file_name} | 幻灯片 {issue.slide_index + 1}", fill="#0f172a", font=small_font)
        draw.text((text_x, y + 25), f"文本: {issue.text[:40]}", fill="#334155", font=font)
        draw.text((text_x, y + 45), f"亮度差 {issue.contrast:.2f} | {issue.reason}", fill="#64748b", font=small_font)

        # 状态条
        bar_color = _contrast_bar_color(issue.text_lum, issue.bg_lum)
        bar_width = int(100 * issue.contrast / 0.5)
        draw.rectangle([width - 130, y + 20, width - 130 + bar_width, y + 40], fill=bar_color)
        draw.rectangle([width - 130, y + 20, width - 30, y + 40], outline="#94a3b8")

        y += row_height

    if not issues:
        draw.text((padding, y + 20), "未发现低对比度文本问题", fill="#15803d", font=font)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(output_path))


def generate_slide_diagnostic_image(
    model: Presentation,
    slide: Slide,
    issues: list[VisibilityIssue],
    output_path: Path,
    width: int = 1280,
) -> None:
    """生成单页幻灯片诊断图：画出形状、文本、背景、低对比标记。"""
    from PIL import Image, ImageDraw, ImageFont

    slide_w = model.slide_width
    slide_h = model.slide_height
    scale = width / slide_w
    height = int(slide_h * scale)

    img = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    # 画背景
    bg_rgb = None
    if slide.background and slide.background.type == "solid":
        bg_rgb = _color_to_rgb(slide.background.color)
    if bg_rgb:
        draw.rectangle([0, 0, width, height], fill=_hex_to_fill(bg_rgb))

    try:
        font = ImageFont.truetype("arial.ttf", 14)
        small_font = ImageFont.truetype("arial.ttf", 11)
    except Exception:
        font = ImageFont.load_default()
        small_font = font

    issue_set = {(i.slide_index, i.shape_name, i.text) for i in issues}

    def _draw_shape(s: Shape, offset_x: int = 0, offset_y: int = 0) -> None:
        x = offset_x + int((s.left or 0) * scale)
        y = offset_y + int((s.top or 0) * scale)
        w = int((s.width or 0) * scale)
        h = int((s.height or 0) * scale)

        fill_rgb = None
        if s.fill and s.fill.type == "solid":
            fill_rgb = _color_to_rgb(s.fill.color)
        if fill_rgb:
            draw.rectangle([x, y, x + w, y + h], fill=_hex_to_fill(fill_rgb), outline="#64748b")
        else:
            draw.rectangle([x, y, x + w, y + h], outline="#64748b")

        text = getattr(s, "text", None)
        if text:
            for para in text.paragraphs:
                for run in para.runs:
                    run_text = run.text.strip()
                    if not run_text:
                        continue
                    has_issue = (slide.index, s.name, run_text) in issue_set
                    text_rgb = _effective_text_rgb(run) or "000000"
                    draw.text((x + 4, y + 4), run_text[:20], fill=_hex_to_fill(text_rgb), font=font)
                    if has_issue:
                        draw.rectangle([x - 2, y - 2, x + w + 2, y + h + 2], outline="#ef4444", width=3)
                    break
                break

        for child in getattr(s, "children", []) or []:
            _draw_shape(child, x, y)

    for shape in slide.shapes:
        _draw_shape(shape)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(output_path))


def scan_all_pptx(
    input_dir: Path,
    output_dir: Path,
    threshold: float = CONTRAST_THRESHOLD,
) -> dict[str, list[VisibilityIssue]]:
    """扫描目录下所有 PPTX，生成可见性报告与诊断图。"""
    from ppt_transfor.parser.presentation import parse_presentation

    output_dir.mkdir(parents=True, exist_ok=True)
    all_issues: dict[str, list[VisibilityIssue]] = {}

    files = sorted(input_dir.glob("*.pptx"))
    files = [f for f in files if not f.name.startswith("~$")]

    for pptx_path in files:
        try:
            model = parse_presentation(str(pptx_path))
        except Exception as e:
            print(f"解析失败 {pptx_path.name}: {e}")
            continue

        issues = detect_presentation_visibility_issues(model, threshold)
        if issues:
            all_issues[pptx_path.name] = issues

        # 生成单页诊断图
        for slide in model.slides:
            slide_issues = [i for i in issues if i.slide_index == slide.index]
            if slide_issues:
                img_path = output_dir / pptx_path.stem / f"slide_{slide.index + 1:03d}.png"
                generate_slide_diagnostic_image(model, slide, slide_issues, img_path)

    # 总报告图
    flat_issues = [i for issues in all_issues.values() for i in issues]
    generate_visibility_report_image(flat_issues, output_dir / "report.png")

    return all_issues
