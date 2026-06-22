"""Chart 图形保留：通过 zip 级后处理把原始 chart 插入渲染后的 PPTX。

python-pptx 不暴露任意 chart 重建能力，因此采用 XML/part 级保留：
- 解析阶段记录原始 <p:graphicFrame> XML 与对应 chart part 路径。
- 渲染阶段跳过 chart，避免生成占位文本框。
- 保存完成后，将原始 chart XML 插入 slide，复制 chart part、chart rels、embedding 等，
  并更新 slide rels 与 [Content_Types].xml。
"""

from __future__ import annotations

import logging
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from os.path import normpath
from pathlib import Path

from lxml import etree

from ppt_transfor.models.schema import Presentation

logger = logging.getLogger(__name__)

# XML 命名空间
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# chart relationship type
RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

# 根据扩展名推断 content type（用于 chart 引用的 embedding、图片等）
_EXT_CONTENT_TYPES: dict[str, str] = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".svg": "image/svg+xml",
}


def _qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


@dataclass
class ChartInfo:
    """单张 chart 的保留信息

    anchor_idx: 在渲染后的 slide 中非 chart 形状列表中的插入锚点位置
                （即排在第 anchor_idx 个非 chart 形状之前）
    """

    chart_xml: str
    chart_part: str | None  # 原始 pptx 中的 chart part 路径，例如 "ppt/charts/chart1.xml"
    anchor_idx: int = 0


def _collect_charts(model: Presentation) -> dict[int, list[ChartInfo]]:
    """按 slide 索引收集需要保留的 chart，并计算每个 chart 应插入的位置"""
    result: dict[int, list[ChartInfo]] = {}
    for slide_idx, slide_model in enumerate(model.slides):
        charts: list[ChartInfo] = []
        non_chart_count = 0
        for shape in slide_model.shapes:
            if shape.shape_type == "chart" and shape.chart_xml:
                charts.append(
                    ChartInfo(
                        chart_xml=shape.chart_xml,
                        chart_part=shape.chart_part,
                        anchor_idx=non_chart_count,
                    )
                )
            else:
                non_chart_count += 1
        if charts:
            result[slide_idx] = charts
    return result


def _max_existing_id(names: list[str], prefix: str, suffix: str) -> int:
    """从形如 prefix{N}.suffix 的文件名中提取最大数字 N"""
    pattern = re.compile(re.escape(prefix) + r"(\d+)" + re.escape(suffix) + r"$")
    max_id = 0
    for name in names:
        m = pattern.match(Path(name).name)
        if m:
            max_id = max(max_id, int(m.group(1)))
    return max_id


def _unique_name(existing: list[str], prefix: str, suffix: str, counter: list[int]) -> str:
    """生成不重复的文件名 prefix{counter}.suffix"""
    while True:
        name = f"{prefix}{counter[0]}{suffix}"
        if name not in [Path(n).name for n in existing]:
            return name
        counter[0] += 1


def _parse_rels_root(data: bytes) -> etree._Element:
    if not data:
        return etree.Element(_qn(NS_REL, "Relationships"), nsmap={None: NS_REL})
    return etree.fromstring(data)


def _serialize_rels(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _next_rid(rels_root: etree._Element) -> str:
    """在现有 relationships 基础上分配新的 rId"""
    max_n = 0
    for rel in rels_root:
        rid = rel.get("Id")
        if rid and rid.startswith("rId"):
            try:
                max_n = max(max_n, int(rid[3:]))
            except ValueError:
                pass
    return f"rId{max_n + 1}"


def _ensure_content_type(zout: zipfile.ZipFile, part_path: str, content_type: str, ct_root: etree._Element):
    """确保 [Content_Types].xml 包含指定 part 的 Override"""
    part_name = f"/{part_path}"
    for override in ct_root.findall(_qn(NS_CT, "Override")):
        if override.get("PartName") == part_name:
            return
    override = etree.SubElement(ct_root, _qn(NS_CT, "Override"))
    override.set("PartName", part_name)
    override.set("ContentType", content_type)


def _content_type_for_part(target_path: str, rel_type: str) -> str | None:
    """根据目标文件扩展名或 relationship type 推断 content type"""
    ext = Path(target_path).suffix.lower()
    if ext in _EXT_CONTENT_TYPES:
        return _EXT_CONTENT_TYPES[ext]
    if rel_type == RT_CHART:
        return "application/vnd.openxmlformats-officedocument.chartml.chart+xml"
    # 其他类型按需要扩展
    return None


def _shape_elements(sp_tree: etree._Element) -> list[etree._Element]:
    """返回 spTree 中代表形状的子元素（排除组合属性等前缀元素）"""
    shape_tags = {
        _qn(NS_P, "sp"),
        _qn(NS_P, "pic"),
        _qn(NS_P, "graphicFrame"),
        _qn(NS_P, "grpSp"),
        _qn(NS_P, "cxnSp"),
    }
    return [child for child in sp_tree if child.tag in shape_tags]


def _max_shape_id(slide_root: etree._Element) -> int:
    """获取 slide XML 中已有形状的最大 cNvPr id"""
    max_id = 0
    for cnvpr in slide_root.iter(_qn(NS_P, "cNvPr")):
        try:
            max_id = max(max_id, int(cnvpr.get("id", 0)))
        except (ValueError, TypeError):
            pass
    return max_id


def _set_shape_id(graphic_frame: etree._Element, new_id: int) -> None:
    """修改 graphicFrame 的 cNvPr id"""
    cnvpr = graphic_frame.find(_qn(NS_P, "nvGraphicFramePr") + "/" + _qn(NS_P, "cNvPr"))
    if cnvpr is not None:
        cnvpr.set("id", str(new_id))


def _copy_chart_dependencies(
    zin: zipfile.ZipFile,
    zout: zipfile.ZipFile,
    orig_chart_part: str,
    new_chart_part: str,
    used_names: list[str],
    emb_counter: list[int],
    ct_root: etree._Element,
) -> None:
    """复制 chart part 本身、其 rels 文件以及 rels 指向的依赖 part"""
    # chart XML 本体
    chart_bytes = zin.read(orig_chart_part)
    zout.writestr(new_chart_part, chart_bytes)
    _ensure_content_type(zout, new_chart_part, "application/vnd.openxmlformats-officedocument.chartml.chart+xml", ct_root)

    orig_chart_id = Path(orig_chart_part).stem  # e.g. chart1
    orig_rels_path = f"ppt/charts/_rels/{orig_chart_id}.xml.rels"
    new_chart_id = Path(new_chart_part).stem
    new_rels_path = f"ppt/charts/_rels/{new_chart_id}.xml.rels"

    rels_data = zin.read(orig_rels_path) if orig_rels_path in zin.namelist() else b""
    rels_root = _parse_rels_root(rels_data)

    for rel in rels_root:
        rel_type = rel.get("Type")
        target_ref = rel.get("Target")
        if not rel_type or not target_ref:
            continue

        # 计算原始依赖 part 在 zip 中的绝对路径
        source_part_path = normpath(str(Path("ppt/charts") / target_ref)).replace("\\", "/")
        if source_part_path not in zin.namelist():
            logger.warning("Chart 依赖 part 在原始 PPTX 中不存在: %s", source_part_path)
            continue

        # 生成目标唯一名称，保留原始依赖所在的文件夹（embeddings / media 等）
        src_name = Path(source_part_path).name
        ext = Path(src_name).suffix
        folder = Path(source_part_path).parent.name  # e.g. embeddings, media
        if ext == ".xlsx":
            new_name = _unique_name(used_names, "Microsoft_Excel_Worksheet", ext, emb_counter)
        else:
            new_name = _unique_name(used_names, "chartDep", ext, emb_counter)

        new_part_path = f"ppt/{folder}/{new_name}"
        # 从 chart 文件位置（ppt/charts/）出发的相对路径
        new_target_ref = f"../{folder}/{new_name}"

        # 复制依赖 part
        dep_bytes = zin.read(source_part_path)
        zout.writestr(new_part_path, dep_bytes)
        used_names.append(new_name)

        # content type
        ct = _content_type_for_part(new_part_path, rel_type)
        if ct:
            _ensure_content_type(zout, new_part_path, ct, ct_root)

        # 更新 relationship 的 Target
        rel.set("Target", new_target_ref)

    zout.writestr(new_rels_path, _serialize_rels(rels_root))


def _process_slide_charts(
    zin: zipfile.ZipFile,
    zout_src: zipfile.ZipFile,
    zout: zipfile.ZipFile,
    slide_idx: int,
    charts: list[ChartInfo],
    used_names: list[str],
    chart_counter: list[int],
    emb_counter: list[int],
    ct_root: etree._Element,
) -> None:
    """处理单个 slide 的所有 chart：插入 XML、复制 part、更新 rels。

    从 zout_src（原始输出 pptx）读取 slide XML/rels，修改后写入 zout（新 pptx），
    避免在 zout 尚未写入 slide 条目时误判为不存在。
    """
    slide_n = slide_idx + 1
    slide_path = f"ppt/slides/slide{slide_n}.xml"
    slide_rels_path = f"ppt/slides/_rels/slide{slide_n}.xml.rels"

    if slide_path not in zout_src.namelist():
        logger.warning("源 PPTX 中不存在 %s，跳过 chart 保留", slide_path)
        return

    slide_xml = etree.fromstring(zout_src.read(slide_path))
    sp_tree = slide_xml.find(f".//{{{NS_P}}}spTree")
    if sp_tree is None:
        logger.warning("%s 中不存在 spTree，跳过 chart 保留", slide_path)
        return

    rels_data = zout_src.read(slide_rels_path) if slide_rels_path in zout_src.namelist() else b""
    rels_root = _parse_rels_root(rels_data)

    shape_elems = _shape_elements(sp_tree)
    next_shape_id = _max_shape_id(slide_xml) + 1

    for chart_info in charts:
        if not chart_info.chart_part:
            logger.warning("chart 缺少 chart_part，无法复制 chart part")
            continue

        # 分配新 chart part 名称
        new_chart_name = _unique_name(used_names, "chart", ".xml", chart_counter)
        new_chart_part = f"ppt/charts/{new_chart_name}"
        used_names.append(new_chart_name)

        # 复制 chart part 及其依赖
        _copy_chart_dependencies(
            zin, zout, chart_info.chart_part, new_chart_part,
            used_names, emb_counter, ct_root,
        )

        # slide -> chart relationship
        slide_rid = _next_rid(rels_root)
        rel = etree.SubElement(rels_root, _qn(NS_REL, "Relationship"))
        rel.set("Id", slide_rid)
        rel.set("Type", RT_CHART)
        rel.set("Target", f"../charts/{new_chart_name}")

        # 解析并调整 chart 的 graphicFrame XML
        try:
            graphic_frame = etree.fromstring(chart_info.chart_xml)
        except Exception as e:
            logger.warning("解析 chart_xml 失败: %s", e)
            continue

        # 更新 <c:chart r:id="..."/>
        chart_el = graphic_frame.find(f".//{{{NS_C}}}chart")
        if chart_el is not None:
            chart_el.set(f"{{{NS_R}}}id", slide_rid)

        # 避免与现有 shape id 冲突
        _set_shape_id(graphic_frame, next_shape_id)
        next_shape_id += 1

        # 插入 slide spTree，保持 z-order
        anchor = chart_info.anchor_idx
        if anchor < len(shape_elems):
            shape_elems[anchor].addprevious(graphic_frame)
        else:
            sp_tree.append(graphic_frame)

    # 写回 slide XML 与 slide rels
    zout.writestr(slide_path, etree.tostring(slide_xml, xml_declaration=True, encoding="UTF-8", standalone=True))
    zout.writestr(slide_rels_path, _serialize_rels(rels_root))


def apply_chart_preservation(
    output_path: str | Path,
    source_path: str | Path,
    model: Presentation,
) -> None:
    """在已渲染的 PPTX 上执行 chart 保留后处理。

    Args:
        output_path: 渲染后的 PPTX 路径（会被覆盖）。
        source_path: 原始 PPTX 路径，用于读取 chart part/embedding。
        model: Presentation 模型，用于定位哪些形状是 chart。
    """
    output_path = Path(output_path)
    source_path = Path(source_path)

    if not source_path.exists():
        logger.debug("原始 PPTX 不存在，跳过 chart 保留: %s", source_path)
        return

    charts_by_slide = _collect_charts(model)
    if not charts_by_slide:
        return

    # 用于生成唯一 part 名的计数器与已占用名集合
    with zipfile.ZipFile(source_path, "r") as zin:
        source_names = zin.namelist()

    used_names = [Path(n).name for n in source_names]
    chart_counter = [_max_existing_id(used_names, "chart", ".xml") + 1]
    emb_counter = [_max_existing_id(
        used_names, "Microsoft_Excel_Worksheet", ".xlsx"
    ) + 1]

    # 重写 zip：复制输出文件，修改/新增 chart 相关条目
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp_file.close()
    tmp_path = tmp_file.name
    try:
        with (
            zipfile.ZipFile(output_path, "r") as zout_src,
            zipfile.ZipFile(source_path, "r") as zin,
            zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout,
        ):
            # 收集会被改写的条目，复制时跳过以避免 zip 中重复
            ct_path = "[Content_Types].xml"
            rewritten_paths = {ct_path}
            for slide_idx in charts_by_slide:
                slide_n = slide_idx + 1
                rewritten_paths.add(f"ppt/slides/slide{slide_n}.xml")
                rewritten_paths.add(f"ppt/slides/_rels/slide{slide_n}.xml.rels")

            # 复制输出文件所有条目
            for info in zout_src.infolist():
                if info.filename in rewritten_paths:
                    continue
                zout.writestr(info, zout_src.read(info.filename))

            # 读取并解析 [Content_Types].xml
            ct_data = zout_src.read(ct_path)
            ct_root = etree.fromstring(ct_data)

            # 逐个 slide 处理 chart
            for slide_idx, charts in charts_by_slide.items():
                _process_slide_charts(
                    zin, zout_src, zout, slide_idx, charts,
                    used_names, chart_counter, emb_counter, ct_root,
                )

            # 写回 [Content_Types].xml
            zout.writestr(
                ct_path,
                etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True),
            )

        # 用修改后的 zip 替换原输出文件
        shutil.move(tmp_path, output_path)
    except Exception:
        Path(tmp_path).unlink(missing_ok=True)
        raise
