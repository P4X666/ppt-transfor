"""视觉对比工具：PPTX → PNG → 像素级 diff。

流程：
1. LibreOffice (soffice) 把 PPTX 转为 PDF
2. PyMuPDF (fitz) 把 PDF 每页转为 PNG
3. Pillow (PIL) 做像素级 diff，生成差异热图

系统依赖：LibreOffice（提供 soffice 命令）。
Windows 下需安装 LibreOffice，本模块会自动查找常见安装路径。
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

# PNG 渲染 DPI（越高越精确，但越慢）
RENDER_DPI = 150
# 像素差异容差（0-255），低于此值视为相同
PIXEL_TOLERANCE = 10


def _find_soffice() -> str | None:
    """查找 LibreOffice 的 soffice 可执行文件路径。

    查找顺序：
    1. PATH 中的 soffice/soffice.exe
    2. Windows 常见安装路径

    Returns:
        soffice 可执行文件路径，找不到返回 None
    """
    # 1. PATH 查找
    soffice = shutil.which("soffice")
    if soffice:
        return soffice

    # 2. Windows 常见安装路径
    if sys.platform == "win32":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for path in candidates:
            if Path(path).exists():
                return path

    return None


def pptx_to_pdf(pptx_path: Path, output_dir: Path) -> Path | None:
    """用 LibreOffice 把 PPTX 转为 PDF。

    Args:
        pptx_path: PPTX 文件路径
        output_dir: PDF 输出目录

    Returns:
        生成的 PDF 文件路径，失败返回 None
    """
    soffice = _find_soffice()
    if soffice is None:
        logger.error(
            "未找到 LibreOffice (soffice)，请安装 LibreOffice 并确保在 PATH 中，"
            "或安装到默认路径 C:\\Program Files\\LibreOffice"
        )
        return None

    output_dir.mkdir(parents=True, exist_ok=True)

    # soffice --headless --convert-to pdf --outdir <output_dir> <pptx>
    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(pptx_path),
    ]
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            logger.error("soffice 转换失败: %s", result.stderr)
            return None
    except subprocess.TimeoutExpired:
        logger.error("soffice 转换超时（120s）")
        return None
    except Exception as e:
        logger.error("soffice 调用异常: %s", e)
        return None

    # PDF 文件名 = PPTX 文件名（换后缀）
    pdf_path = output_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        logger.error("PDF 未生成: %s", pdf_path)
        return None

    return pdf_path


def pdf_to_images(pdf_path: Path, output_dir: Path, dpi: int = RENDER_DPI) -> list[Path]:
    """用 PyMuPDF 把 PDF 每页转为 PNG。

    Args:
        pdf_path: PDF 文件路径
        output_dir: PNG 输出目录
        dpi: 渲染 DPI

    Returns:
        PNG 文件路径列表（按页序）
    """
    import fitz  # PyMuPDF

    output_dir.mkdir(parents=True, exist_ok=True)
    images: list[Path] = []

    doc = fitz.open(str(pdf_path))
    try:
        # 缩放因子：DPI / 72（PDF 默认 72 DPI）
        zoom = dpi / 72
        matrix = fitz.Matrix(zoom, zoom)

        for page_idx in range(len(doc)):
            page = doc[page_idx]
            pixmap = page.get_pixmap(matrix=matrix)
            img_path = output_dir / f"page_{page_idx:03d}.png"
            pixmap.save(str(img_path))
            images.append(img_path)
    finally:
        doc.close()

    return images


def pptx_to_images(pptx_path: Path, output_dir: Path) -> list[Path]:
    """PPTX → PDF（soffice）→ PNG 每页（PyMuPDF）。

    Args:
        pptx_path: PPTX 文件路径
        output_dir: 临时文件输出目录

    Returns:
        PNG 文件路径列表（按页序），失败返回空列表
    """
    pdf_path = pptx_to_pdf(pptx_path, output_dir)
    if pdf_path is None:
        return []
    return pdf_to_images(pdf_path, output_dir)


def compare_images(img1: Path, img2: Path, tolerance: int = PIXEL_TOLERANCE) -> dict:
    """像素级对比两张图片，返回差异度。

    Args:
        img1: 原始图片路径
        img2: 转换后图片路径
        tolerance: 像素差异容差（0-255）

    Returns:
        {"diff_ratio": float, "max_diff": int, "size_match": bool}
        - diff_ratio: 差异像素占比（0.0-1.0）
        - max_diff: 最大像素差异（0-255）
        - size_match: 两图尺寸是否一致
    """
    from PIL import Image, ImageChops

    im1 = Image.open(str(img1)).convert("RGB")
    im2 = Image.open(str(img2)).convert("RGB")

    # 尺寸不一致时，以较小的为准裁剪
    size_match = im1.size == im2.size
    if not size_match:
        min_w = min(im1.width, im2.width)
        min_h = min(im1.height, im2.height)
        im1 = im1.crop((0, 0, min_w, min_h))
        im2 = im2.crop((0, 0, min_w, min_h))

    diff = ImageChops.difference(im1, im2)

    # 统计差异像素
    diff_data = list(diff.getdata())
    total_pixels = len(diff_data)
    diff_pixels = 0
    max_diff = 0

    for r, g, b in diff_data:
        pixel_diff = max(r, g, b)
        if pixel_diff > max_diff:
            max_diff = pixel_diff
        if pixel_diff > tolerance:
            diff_pixels += 1

    diff_ratio = diff_pixels / total_pixels if total_pixels > 0 else 0.0

    return {
        "diff_ratio": diff_ratio,
        "max_diff": max_diff,
        "size_match": size_match,
    }


def generate_diff_heatmap(img1: Path, img2: Path, output: Path, tolerance: int = PIXEL_TOLERANCE) -> None:
    """生成差异热图（差异区域红色高亮叠加）。

    Args:
        img1: 原始图片路径
        img2: 转换后图片路径
        output: 热图输出路径
        tolerance: 像素差异容差
    """
    from PIL import Image, ImageChops

    im1 = Image.open(str(img1)).convert("RGB")
    im2 = Image.open(str(img2)).convert("RGB")

    # 尺寸不一致时裁剪
    if im1.size != im2.size:
        min_w = min(im1.width, im2.width)
        min_h = min(im1.height, im2.height)
        im1 = im1.crop((0, 0, min_w, min_h))
        im2 = im2.crop((0, 0, min_w, min_h))

    diff = ImageChops.difference(im1, im2)

    # 生成热图：差异区域红色高亮，无差异区域保留原图（变暗）
    heatmap = Image.new("RGB", im1.size)
    heatmap_data = []
    for (r, g, b), orig in zip(diff.getdata(), im1.getdata()):
        pixel_diff = max(r, g, b)
        if pixel_diff > tolerance:
            # 差异区域：红色高亮，强度按差异大小
            intensity = min(255, pixel_diff * 3)
            heatmap_data.append((intensity, 0, 0))
        else:
            # 无差异区域：原图变暗（便于聚焦差异）
            heatmap_data.append((orig[0] // 3, orig[1] // 3, orig[2] // 3))

    heatmap.putdata(heatmap_data)
    output.parent.mkdir(parents=True, exist_ok=True)
    heatmap.save(str(output))


def visual_compare(
    original_pptx: Path,
    converted_pptx: Path,
    output_dir: Path,
) -> dict:
    """对比两个 PPTX 的视觉差异，逐页对比。

    Args:
        original_pptx: 原始 PPTX 路径
        converted_pptx: 转换后 PPTX 路径
        output_dir: 输出目录（热图、报告）

    Returns:
        {
            "pages": [{"page": int, "diff_ratio": float, "max_diff": int, "size_match": bool}, ...],
            "avg_diff_ratio": float,
            "max_page_diff": float,
            "heatmap_paths": [Path, ...],
            "success": bool,
            "error": str | None,
        }
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. 两个 PPTX 分别转 PNG
    orig_img_dir = output_dir / "orig_images"
    conv_img_dir = output_dir / "conv_images"

    orig_images = pptx_to_images(original_pptx, orig_img_dir)
    if not orig_images:
        return {
            "pages": [],
            "avg_diff_ratio": 0.0,
            "max_page_diff": 0.0,
            "heatmap_paths": [],
            "success": False,
            "error": "原始 PPTX 转 PNG 失败（请检查 LibreOffice 是否安装）",
        }

    conv_images = pptx_to_images(converted_pptx, conv_img_dir)
    if not conv_images:
        return {
            "pages": [],
            "avg_diff_ratio": 0.0,
            "max_page_diff": 0.0,
            "heatmap_paths": [],
            "success": False,
            "error": "转换后 PPTX 转 PNG 失败（请检查 LibreOffice 是否安装）",
        }

    # 2. 逐页对比
    pages = []
    heatmap_paths = []
    max_page_diff = 0.0
    total_ratio = 0.0
    compare_count = 0

    for page_idx in range(max(len(orig_images), len(conv_images))):
        orig_img = orig_images[page_idx] if page_idx < len(orig_images) else None
        conv_img = conv_images[page_idx] if page_idx < len(conv_images) else None

        if orig_img is None or conv_img is None:
            pages.append({
                "page": page_idx,
                "diff_ratio": 1.0,
                "max_diff": 255,
                "size_match": False,
                "note": "页数不匹配",
            })
            max_page_diff = 1.0
            total_ratio += 1.0
            compare_count += 1
            continue

        result = compare_images(orig_img, conv_img)
        pages.append({
            "page": page_idx,
            "diff_ratio": result["diff_ratio"],
            "max_diff": result["max_diff"],
            "size_match": result["size_match"],
        })

        if result["diff_ratio"] > 0.01:
            # 差异大于 1% 才生成热图
            heatmap_path = output_dir / f"diff_page_{page_idx:03d}.png"
            generate_diff_heatmap(orig_img, conv_img, heatmap_path)
            heatmap_paths.append(heatmap_path)

        total_ratio += result["diff_ratio"]
        compare_count += 1
        if result["diff_ratio"] > max_page_diff:
            max_page_diff = result["diff_ratio"]

    avg_diff_ratio = total_ratio / compare_count if compare_count > 0 else 0.0

    return {
        "pages": pages,
        "avg_diff_ratio": avg_diff_ratio,
        "max_page_diff": max_page_diff,
        "heatmap_paths": heatmap_paths,
        "success": True,
        "error": None,
    }
