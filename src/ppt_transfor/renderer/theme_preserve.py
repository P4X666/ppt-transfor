"""主题保留：通过 zip 级后处理把原始 PPT 的主题和表格样式复制到转换后的 PPT。

python-pptx 创建新 Presentation 时使用默认 Office 主题（accent1=4F81BD 蓝），
导致所有 schemeClr 引用映射到错误颜色。此模块在保存后替换所有 theme*.xml 和
tableStyles.xml，确保转换后 PPT 使用与原始 PPT 相同的主题颜色。

注意：复制 theme2.xml 等新增 part 时，必须同步在 [Content_Types].xml 中添加
Override 声明，否则 Office 会判定文件损坏无法打开。
"""

from __future__ import annotations

import logging
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

logger = logging.getLogger(__name__)

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
THEME_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml"
TABLE_STYLES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"
)


def _qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _collect_source_overrides(
    source_names: list[str],
    ct_data: bytes,
) -> dict[str, str]:
    """从原始 PPT 的 [Content_Types].xml 提取需要同步的 Override。

    返回 {part_path: content_type}，仅包含 theme*.xml 和 tableStyles.xml。
    """
    overrides: dict[str, str] = {}
    try:
        ct_root = etree.fromstring(ct_data)
    except Exception:
        return overrides

    for override in ct_root.findall(_qn(NS_CT, "Override")):
        part_name = override.get("PartName", "")
        content_type = override.get("ContentType", "")
        # PartName 以 / 开头，去掉前缀得到 zip 内路径
        part_path = part_name.lstrip("/")
        is_theme = re.match(r"ppt/theme/theme\d+\.xml$", part_path) is not None
        is_table_styles = part_path == "ppt/tableStyles.xml"
        if (is_theme or is_table_styles) and content_type:
            overrides[part_path] = content_type
    return overrides


def _scan_referenced_themes(zip_reader: zipfile.ZipFile) -> set[str]:
    """扫描输出 PPT 的所有 .rels，找出实际被引用的 theme 文件路径。

    只复制被引用的 theme 文件，避免「游离 part」导致 Office 判定文件损坏。
    典型场景：原始 PPT 的 notesMaster 引用 theme2.xml，转换后 notesMaster
    丢失，theme2.xml 成了无主 part，Office 会因结构异常报损坏。

    注意：.rels 中的 Target 是相对路径（如 "../theme/theme1.xml" 或
    "theme/theme1.xml"），不是完整的 "ppt/theme/theme1.xml"。此函数
    将相对 Target 归一化为 "ppt/theme/themeN.xml" 形式。

    Args:
        zip_reader: 输出 PPT 的 ZipFile 对象

    Returns:
        被引用的 theme 文件路径集合，如 {"ppt/theme/theme1.xml"}
    """
    referenced: set[str] = set()
    # 匹配 Target 中以 theme/themeN.xml 结尾的路径（相对路径）
    # 如 "../theme/theme1.xml"、"theme/theme1.xml"
    theme_pattern = re.compile(r'Target="[^"]*?(theme/theme\d+\.xml)"')
    for info in zip_reader.infolist():
        if not info.filename.endswith(".rels"):
            continue
        try:
            data = zip_reader.read(info.filename).decode("utf-8", errors="ignore")
        except Exception:
            continue
        for m in theme_pattern.finditer(data):
            # 归一化为 ppt/theme/themeN.xml
            theme_rel = m.group(1)
            referenced.add(f"ppt/{theme_rel}")
    return referenced


def _ensure_overrides(
    ct_root: etree._Element,
    overrides_to_add: dict[str, str],
) -> None:
    """确保 [Content_Types].xml 包含指定 part 的 Override，缺失则补齐。"""
    existing = {
        o.get("PartName", "").lstrip("/") for o in ct_root.findall(_qn(NS_CT, "Override"))
    }
    for part_path, content_type in overrides_to_add.items():
        if part_path in existing:
            continue
        override = etree.SubElement(ct_root, _qn(NS_CT, "Override"))
        override.set("PartName", f"/{part_path}")
        override.set("ContentType", content_type)


def apply_theme_preservation(output_path: str | Path, source_path: str | Path) -> None:
    """将原始 PPT 的主题和表格样式复制到转换后的 PPT。

    动态匹配所有 ppt/theme/theme*.xml（包括 theme1.xml、theme2.xml 等），
    避免因 theme2.xml 丢失导致 chart 颜色渲染异常。

    复制新增的 theme part 时，同步在 [Content_Types].xml 中补充 Override，
    避免 Office 因缺少 content type 声明而判定文件损坏。

    Args:
        output_path: 转换后的 PPTX 路径（会被覆盖）
        source_path: 原始 PPTX 路径
    """
    output_path = Path(output_path)
    source_path = Path(source_path)

    if not source_path.exists():
        logger.debug("原始 PPTX 不存在，跳过主题保留: %s", source_path)
        return

    # 先扫描输出 PPT 的 .rels，找出实际被引用的 theme 文件
    # 只复制被引用的 theme，避免「游离 part」导致 Office 判定文件损坏
    # 典型场景：原始 PPT 的 notesMaster 引用 theme2.xml，转换后 notesMaster
    # 丢失，若仍复制 theme2.xml 会变成无主 part
    with zipfile.ZipFile(output_path, "r") as zout_scan:
        referenced_themes = _scan_referenced_themes(zout_scan)
    logger.debug("输出 PPT 实际引用的 theme: %s", sorted(referenced_themes))

    # 动态扫描原始 PPT 中所有 theme*.xml 和 tableStyles.xml
    with zipfile.ZipFile(source_path, "r") as zin:
        source_names = zin.namelist()
        # 只复制「原始 PPT 有 且 输出 PPT 实际引用」的 theme 文件
        all_source_themes = [
            n for n in source_names
            if n.startswith("ppt/theme/theme") and n.endswith(".xml")
        ]
        parts_to_copy = [n for n in all_source_themes if n in referenced_themes]
        if "ppt/tableStyles.xml" in source_names:
            parts_to_copy.append("ppt/tableStyles.xml")

        if not parts_to_copy:
            return

        # 读取需要复制的文件内容
        parts_data = {p: zin.read(p) for p in parts_to_copy}
        # 读取原始 PPT 的 [Content_Types].xml，提取需要同步的 Override
        source_ct_data = zin.read("[Content_Types].xml")
        source_overrides = _collect_source_overrides(source_names, source_ct_data)

    # 重写 zip：复制输出文件所有条目，替换主题相关条目
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp_file.close()
    tmp_path = tmp_file.name

    try:
        with (
            zipfile.ZipFile(output_path, "r") as zout_src,
            zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            ct_path = "[Content_Types].xml"
            ct_data = zout_src.read(ct_path)
            ct_root = etree.fromstring(ct_data)

            # 仅对「输出中缺失但原始中有 且 实际被引用」的 part 补 Override
            output_existing = {
                o.get("PartName", "").lstrip("/")
                for o in ct_root.findall(_qn(NS_CT, "Override"))
            }
            overrides_to_add = {
                p: t for p, t in source_overrides.items()
                if p not in output_existing and p in parts_data
            }
            if overrides_to_add:
                _ensure_overrides(ct_root, overrides_to_add)
                logger.debug("补充 Content_Types Override: %s", list(overrides_to_add))

            new_ct_data = etree.tostring(
                ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

            # 复制输出文件所有条目，跳过要替换的条目与 Content_Types
            for info in zout_src.infolist():
                if info.filename in parts_data or info.filename == ct_path:
                    continue
                zout.writestr(info, zout_src.read(info.filename))

            # 写入更新后的 [Content_Types].xml
            zout.writestr(ct_path, new_ct_data)

            # 写入原始 PPT 的主题和表格样式
            for part_path, data in parts_data.items():
                zout.writestr(part_path, data)

        shutil.move(tmp_path, output_path)
    except Exception:
        Path(tmp_path).unlink(missing_ok=True)
        raise
