"""JSON schema 数据模型定义。

所有几何尺寸（left/top/width/height/line.width/space_before/space_after/font.size）
统一存 EMU 整数，避免浮点转换误差；line_spacing 为行距倍数（float）。
颜色保留原始来源类型（rgb/theme/scheme），确保高保真回写。
"""

from __future__ import annotations

from typing import Literal, Optional

from pydantic import BaseModel, ConfigDict


class Color(BaseModel):
    """颜色：保留原始来源以便高保真回写。"""

    model_config = ConfigDict(extra="allow")

    type: Literal["rgb", "theme", "scheme"]
    value: str


class Font(BaseModel):
    """字体属性，size 为 EMU 整数（44pt = 558800 EMU）。"""

    model_config = ConfigDict(extra="allow")

    name: Optional[str] = None
    size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    cap: Optional[str] = None  # "all" 表示全大写，"small" 表示小型大写
    color: Optional[Color] = None


class Run(BaseModel):
    """文本 run：一段同格式文本。"""

    model_config = ConfigDict(extra="allow")

    text: str
    font: Font = Font()


class Paragraph(BaseModel):
    """段落：含对齐、层级、间距、行距、runs。"""

    model_config = ConfigDict(extra="allow")

    alignment: Optional[str] = None
    level: int = 0
    space_before: Optional[int] = None
    space_after: Optional[int] = None
    line_spacing: Optional[float] = None
    runs: list[Run] = []


class Text(BaseModel):
    """文本框内容：含自动换行、自适应、垂直对齐、边距、段落列表。"""

    model_config = ConfigDict(extra="allow")

    word_wrap: Optional[bool] = None
    auto_size: Optional[str] = None
    vertical_anchor: Optional[str] = None
    # 文本框内部边距（EMU），来自 <a:bodyPr> 的 lIns/tIns/rIns/bIns
    margin_left: Optional[int] = None
    margin_top: Optional[int] = None
    margin_right: Optional[int] = None
    margin_bottom: Optional[int] = None
    paragraphs: list[Paragraph] = []


class GradientStop(BaseModel):
    """渐变停止点：位置 0.0~1.0 与对应颜色。"""

    model_config = ConfigDict(extra="allow")

    position: float
    color: Color


class Fill(BaseModel):
    """填充：纯色/渐变/图案/图片/无。"""

    model_config = ConfigDict(extra="allow")

    type: str = "none"
    color: Optional[Color] = None
    # 渐变相关属性
    gradient_type: Optional[str] = None  # linear / radial / rect / path
    gradient_angle: Optional[float] = None  # 线性渐变角度（EMU，如 5400000）
    gradient_stops: list[GradientStop] = []


class Line(BaseModel):
    """边框线条。"""

    model_config = ConfigDict(extra="allow")

    width: Optional[int] = None
    color: Optional[Color] = None
    dash: Optional[str] = None
    # 箭头线端点类型（如 "triangle"），对应 OpenXML <a:headEnd>/<a:tailEnd> 的 type 属性
    head_arrow_type: Optional[str] = None
    tail_arrow_type: Optional[str] = None
    # 线条无填充标记：原始 XML 中 <a:ln> 有宽度但无 solidFill 子元素
    # 渲染时需显式设置 noFill，避免继承 p:style 或主题默认色（蓝色）
    no_fill: bool = False


class Shadow(BaseModel):
    """阴影。"""

    model_config = ConfigDict(extra="allow")

    enabled: bool = False


class Crop(BaseModel):
    """图片裁剪比例（0.0~1.0）。"""

    model_config = ConfigDict(extra="allow")

    left: float = 0.0
    top: float = 0.0
    right: float = 0.0
    bottom: float = 0.0


class CellBorder(BaseModel):
    """单元格边框：保留原始边框信息（宽度/颜色/无填充）。"""

    model_config = ConfigDict(extra="allow")

    width: Optional[int] = None
    color: Optional[Color] = None
    # 无填充标记：原始 XML 中 <a:lnL>/<a:lnR> 等有 width 但无 solidFill
    no_fill: bool = False


class TableCell(BaseModel):
    """表格单元格。"""

    model_config = ConfigDict(extra="allow")

    text: Text = Text()
    fill: Optional[Fill] = None
    span_x: int = 1
    span_y: int = 1
    # 单元格四边边框，对应 <a:tcPr> 下的 lnL/lnR/lnT/lnB
    border_left: Optional[CellBorder] = None
    border_right: Optional[CellBorder] = None
    border_top: Optional[CellBorder] = None
    border_bottom: Optional[CellBorder] = None


class Table(BaseModel):
    """表格：行数、列数、单元格矩阵、首行表头标记。"""

    model_config = ConfigDict(extra="allow")

    rows: int
    cols: int
    cells: list[list[TableCell]] = []
    first_row_header: bool = False
    # 表格样式 ID（OOXML tableStyleId GUID），保留原始样式引用
    style_id: Optional[str] = None
    # tblPr 属性：特殊行列样式标记
    first_col: bool = False
    last_row: bool = False
    band_row: bool = False
    # 列宽（EMU 整数），保留原始列宽避免 add_table 默认均分
    col_widths: list[int] = []


class Shape(BaseModel):
    """形状统一模型：通过 shape_type 区分类型，类型特有字段可选。"""

    model_config = ConfigDict(extra="allow")

    shape_id: Optional[str] = None
    name: Optional[str] = None
    shape_type: str = "unknown"
    left: Optional[int] = None
    top: Optional[int] = None
    width: Optional[int] = None
    height: Optional[int] = None
    rotation: float = 0.0

    fill: Optional[Fill] = None
    line: Optional[Line] = None
    shadow: Optional[Shadow] = None

    # 是否有 <p:style> 元素（主题样式引用）
    # add_shape 会自动创建 p:style，但原始 PPT 可能没有
    # 渲染时若 has_style=False，移除 add_shape 创建的 p:style，避免主题边框干扰
    has_style: bool = False
    # p:style 的完整 XML（若原始 PPT 有 p:style），用于回写
    # add_textbox 不创建 p:style，需要从原始 PPT 复制
    style_xml: Optional[str] = None

    # 文本框/自选图形/占位符共有
    text: Optional[Text] = None

    # 自选图形特有
    auto_shape_type: Optional[str] = None
    adjustments: list[float] = []

    # 图片特有
    # data_base64 保留向后兼容旧 JSON；新解析流程改用 image_path 引用外部文件
    data_base64: Optional[str] = None
    # 图片相对路径（如 "media/abc123.png"），相对 out/ 目录
    image_path: Optional[str] = None
    image_format: Optional[str] = None
    crop: Optional[Crop] = None

    # 表格特有
    table: Optional[Table] = None

    # 图表特有：保留原始 <p:graphicFrame> XML，渲染时可选转为图片或 XML 级保留
    chart_xml: Optional[str] = None
    # 原始 chart part 路径，例如 "ppt/charts/chart1.xml"
    chart_part: Optional[str] = None

    # 组合特有
    children: list[Shape] = []
    # 组合特有：子坐标系（chOff/chExt），用于正确渲染子形状坐标空间
    child_offset: Optional[tuple[int, int]] = None  # (x, y) EMU
    child_extent: Optional[tuple[int, int]] = None  # (cx, cy) EMU

    # 连接线特有
    begin_x: Optional[int] = None
    begin_y: Optional[int] = None
    end_x: Optional[int] = None
    end_y: Optional[int] = None


class Background(BaseModel):
    """幻灯片背景。"""

    model_config = ConfigDict(extra="allow")

    type: str = "none"
    color: Optional[Color] = None


class Slide(BaseModel):
    """单页幻灯片。"""

    model_config = ConfigDict(extra="allow")

    index: int
    layout_name: str = "Blank"
    background: Optional[Background] = None
    shapes: list[Shape] = []


class Presentation(BaseModel):
    """PPT 文档根模型。"""

    model_config = ConfigDict(extra="allow")

    version: str = "1.0"
    source_file: str
    slide_width: int
    slide_height: int
    slides: list[Slide] = []
