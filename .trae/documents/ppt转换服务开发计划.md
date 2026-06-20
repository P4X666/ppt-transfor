# PPT 转换服务开发计划

## 概述

开发一个 Python 后端服务，基于 `python-pptx` 实现 PPT 与 JSON 的双向转换：
- **功能1**：PPT → JSON（解析）
- **功能2**：JSON → PPT（渲染）
- **功能3**：通过「原始PPT → JSON → 转换后PPT → JSON」往返对比，定位解析/渲染代码缺陷

包管理器：`uv`；虚拟环境：`uv venv`；核心依赖：`python-pptx`。

---

## 当前状态分析

### 项目现状
- 工作目录 `i:\wefor\ppt-transfor` 仅有 `input/` 目录，无任何 Python 代码
- `input/` 下有 17 个 `.pptx` 测试文件，按页面类型命名：
  - `cover_ending_page.pptx`：封面/封底
  - `index_page.pptx`：目录页
  - `section_page.pptx`：章节页
  - `content_page_component_*.pptx`：内容页各类组件（card / timeline / calendar / closed_loop / imgtext / key_dates / ladder / multi_ring / multicolumn / radial_list / step / table / text / vn）
- 环境已就绪：`uv 0.11.19` + `Python 3.13.2`

### 设计决策（已与用户确认）
1. **对比方式**：结构化 JSON 对比（原始JSON vs 转换后JSON 的字段级 diff）
2. **元素范围**：全量支持（文本框/自选图形/图片/表格/组合/连接线 + 通用属性）

---

## 项目结构

```
ppt-transfor/
├── input/                          # 测试用PPT（已存在）
├── out/                            # 输出目录（运行时自动创建）
│   ├── json/                       # 原始PPT解析后的JSON
│   ├── pptx/                       # JSON重建后的PPT
│   └── compare/                    # 对比报告
├── src/
│   └── ppt_transfor/
│       ├── __init__.py
│       ├── models/                 # 数据模型（pydantic）
│       │   ├── __init__.py
│       │   └── schema.py           # JSON schema 定义（Presentation/Slide/Shape/Text/Table/Image/Group...）
│       ├── parser/                 # PPT → JSON
│       │   ├── __init__.py
│       │   ├── presentation.py     # 入口：遍历 slides，提取尺寸/主题
│       │   ├── slide.py            # 单页：背景 + shapes 列表
│       │   ├── shape.py            # 形状通用属性：位置/旋转/填充/边框/阴影 + 类型分发
│       │   ├── text.py             # 文本框：paragraphs/runs/字体/颜色/对齐/间距
│       │   ├── autoshape.py        # 自选图形：auto_shape_type/adjustments/几何
│       │   ├── table.py            # 表格：单元格/合并/样式/文本
│       │   ├── image.py            # 图片：base64/格式/裁剪
│       │   ├── group.py            # 组合：递归子形状
│       │   └── connector.py        # 连接线：起终点/样式
│       ├── renderer/               # JSON → PPT
│       │   ├── __init__.py
│       │   ├── presentation.py     # 创建 Presentation，设置尺寸，遍历 slides
│       │   ├── slide.py            # 添加空白布局页，设置背景
│       │   ├── shape.py            # 通用属性回写 + 类型分发
│       │   ├── text.py             # 文本框重建
│       │   ├── autoshape.py        # 自选图形重建
│       │   ├── table.py            # 表格重建
│       │   ├── image.py            # 图片重建（base64 → blob）
│       │   ├── group.py            # 组合重建（递归）
│       │   └── connector.py        # 连接线重建
│       ├── comparator/             # JSON 对比
│       │   ├── __init__.py
│       │   └── differ.py           # 深度 diff，输出差异路径+旧值+新值
│       └── utils/
│           ├── __init__.py
│           ├── units.py            # EMU ↔ pt/cm 换算（JSON 内部统一存 EMU 整数）
│           ├── color.py            # 颜色解析：RGB / 主题色 / scheme 色
│           └── xml_helper.py       # 直接操作 lxml 处理 python-pptx 未暴露的属性
├── tests/
│   └── test_roundtrip.py           # 往返测试：input 下每个 pptx 跑 roundtrip 并断言无差异
├── main.py                         # CLI 入口（click）
├── pyproject.toml                  # uv 项目配置 + 依赖
└── .gitignore                      # 忽略 .venv / out / __pycache__
```

---

## JSON Schema 设计

```jsonc
{
  "version": "1.0",
  "source_file": "cover_ending_page.pptx",
  "slide_width": 9144000,          // EMU
  "slide_height": 6858000,         // EMU
  "slides": [
    {
      "index": 0,
      "layout_name": "Blank",
      "background": {              // 可选
        "type": "solid",           // solid | gradient | image | none
        "color": { "type": "rgb", "value": "FFFFFF" }
      },
      "shapes": [
        {
          "shape_id": "sp1",
          "name": "Title 1",
          "shape_type": "text_box",  // text_box | auto_shape | picture | table | group | connector | placeholder
          "left": 457200, "top": 228600, "width": 8229600, "height": 1143000,  // EMU
          "rotation": 0,
          "fill": { "type": "solid", "color": { "type": "rgb", "value": "FF0000" } }
                  // solid | gradient | pattern | picture | none
          "line": { "width": 12700, "color": {...}, "dash": "solid" },
          "shadow": { "enabled": false },
          // 类型特有字段
          "text": {
            "word_wrap": true,
            "auto_size": "none",
            "vertical_anchor": "top",
            "paragraphs": [
              {
                "alignment": "center",
                "level": 0,
                "space_before": 0, "space_after": 0, "line_spacing": 1.0,
                "runs": [
                  {
                    "text": "标题",
                    "font": {
                      "name": "微软雅黑", "size": 44, "bold": false, "italic": false,
                      "underline": false, "color": { "type": "rgb", "value": "000000" }
                    }
                  }
                ]
              }
            ]
          }
        }
      ]
    }
  ]
}
```

### 关键字段说明
- **单位**：所有尺寸统一存 EMU 整数，避免浮点误差
- **颜色**：`{ "type": "rgb"|"theme"|"scheme", "value": "..." }`，保留原始来源以便高保真回写
- **图片**：`{ "data_base64": "...", "format": "png", "crop": {...} }`，base64 自包含
- **表格**：`{ "rows": [...], "cols": [...], "cells": [[{...}]], "first_row_header": true }`
- **组合**：`{ "children": [<shape>...] }` 递归
- **自选图形**：`{ "auto_shape_type": "ROUNDED_RECTANGLE", "adjustments": [0.1] }`

---

## 实施步骤

### 步骤 1：项目初始化
**文件**：`pyproject.toml`、`.gitignore`、`src/ppt_transfor/__init__.py`

- 用 `uv init` 初始化项目，生成 `pyproject.toml`
- 添加依赖：`uv add python-pptx pydantic click rich`
- 创建虚拟环境：`uv venv`（uv 自动管理）
- 配置 `.gitignore`：`.venv/`、`out/`、`__pycache__/`、`*.pyc`
- 建立目录骨架（上述结构中的所有 `__init__.py`）

### 步骤 2：数据模型层
**文件**：`src/ppt_transfor/models/schema.py`

- 用 pydantic 定义所有 schema 类：`Presentation`、`Slide`、`Shape`（基类）、`TextBox`、`AutoShape`、`Picture`、`Table`、`Group`、`Connector`、`Fill`、`Line`、`Shadow`、`Text`、`Paragraph`、`Run`、`Font`、`Color`
- 提供 `model_dump_json(exclude_none=True)` 序列化与 `model_validate_json` 反序列化
- 字段命名与上述 JSON schema 一致

### 步骤 3：工具层
**文件**：`src/ppt_transfor/utils/{units,color,xml_helper}.py`

- `units.py`：`emu_to_pt` / `pt_to_emu` / `emu_to_cm` 等换算函数
- `color.py`：解析 `RGBColor`、主题色（`theme_color`）、scheme 色，统一输出 `Color` 模型
- `xml_helper.py`：封装 lxml 直接读写 `shape._element` 的辅助函数，用于处理 python-pptx 未暴露的属性（如阴影、渐变、自定义几何）

### 步骤 4：解析器（PPT → JSON）
**文件**：`src/ppt_transfor/parser/*.py`

按依赖顺序实现：
1. `text.py`：`parse_text_frame(tf) -> Text`，遍历段落与 run，提取字体/颜色/对齐/间距
2. `image.py`：`parse_picture(shape) -> Picture`，读取 `shape.image.blob` 转 base64，记录格式与裁剪
3. `table.py`：`parse_table(shape) -> Table`，遍历单元格，记录合并（`span_x`/`span_y`）与文本
4. `autoshape.py`：`parse_autoshape(shape) -> AutoShape`，记录 `auto_shape_type` 与 `adjustments`
5. `group.py`：`parse_group(shape) -> Group`，递归调用 `shape.py` 的 `parse_shape`
6. `connector.py`：`parse_connector(shape) -> Connector`，记录起终点与样式
7. `shape.py`：`parse_shape(shape) -> Shape`，提取通用属性（位置/旋转/填充/边框/阴影），按 `shape.shape_type` 分发到上述解析器
8. `slide.py`：`parse_slide(slide) -> Slide`，解析背景 + 遍历 `slide.shapes`
9. `presentation.py`：`parse_presentation(path) -> Presentation`，打开 pptx，读取尺寸，遍历 slides

### 步骤 5：渲染器（JSON → PPT）
**文件**：`src/ppt_transfor/renderer/*.py`

与解析器对称实现：
1. `text.py`：`render_text_box(slide, shape_model) -> Shape`，创建文本框，回写段落/run/字体
2. `image.py`：`render_picture(slide, shape_model) -> Shape`，base64 解码为 blob，`slide.shapes.add_picture`
3. `table.py`：`render_table(slide, shape_model) -> Shape`，`add_table` 后填充单元格，处理合并
4. `autoshape.py`：`render_autoshape(slide, shape_model) -> Shape`，`add_shape(MSO_SHAPE.XXX)`，回写 adjustments
5. `group.py`：`render_group(slide, shape_model) -> Shape`，`group_shapes` 递归
6. `connector.py`：`render_connector(slide, shape_model) -> Shape`
7. `shape.py`：`render_shape(slide, shape_model) -> Shape`，按 `shape_type` 分发，回写通用属性
8. `slide.py`：`render_slide(prs, slide_model) -> Slide`，`prs.slides.add_slide(blank_layout)`，设置背景
9. `presentation.py`：`render_presentation(model) -> Presentation`，创建 `Presentation()`，设置尺寸，遍历 slides，`prs.save(path)`

### 步骤 6：对比器
**文件**：`src/ppt_transfor/comparator/differ.py`

- `diff_json(original: dict, converted: dict) -> list[DiffItem]`
- 深度递归对比两个 dict，记录差异：`{ "path": "slides[0].shapes[2].text.paragraphs[0].runs[1].font.size", "original": 44, "converted": null }`
- **忽略列表**（已知往返必然变化的字段）：
  - `shape_id`（重建后 ID 会变）
  - `name`（部分自动命名会变）
  - `layout_name`（统一用 Blank）
- **浮点容差**：浮点数差异 < 1e-6 视为相等
- 输出差异列表 + 统计摘要（总差异数、按类型分组）

### 步骤 7：CLI 入口
**文件**：`main.py`

使用 click 实现命令：
```bash
# 解析单个PPT
uv run python main.py parse <pptx_path> [--out out/json/xxx.json]

# 解析 input/ 下所有PPT
uv run python main.py parse-all

# 从JSON渲染PPT
uv run python main.py render <json_path> [--out out/pptx/xxx.pptx]

# 渲染 out/json/ 下所有JSON
uv run python main.py render-all

# 对比原始PPT与转换后PPT（内部：两者都解析为JSON后diff）
uv run python main.py compare <original_pptx> <converted_pptx> [--report out/compare/xxx.txt]

# 完整往返：解析→渲染→对比，输出报告
uv run python main.py roundtrip <pptx_path>

# 对 input/ 下所有PPT执行往返测试
uv run python main.py roundtrip-all
```

### 步骤 8：往返测试
**文件**：`tests/test_roundtrip.py`

- 用 pytest 参数化，对 `input/` 下每个 pptx 执行：解析→渲染→解析→对比
- 断言差异列表为空（或在可接受范围内）
- 首次运行预期会有差异，用于暴露 parser/renderer 缺陷，逐步修复

---

## 关键技术要点

### 1. 单位处理
python-pptx 内部用 EMU（914400 EMU = 1 英寸 = 2.54cm）。JSON 统一存 EMU 整数，避免浮点误差。

### 2. 颜色保真
python-pptx 颜色来源多样：
- `RGBColor`（直接 RGB）
- 主题色（`MSO_THEME_COLOR.ACCENT1` 等）
- scheme 色

解析时保留原始类型：`{ "type": "theme", "value": "ACCENT1" }`，渲染时按类型回写，避免主题色被错误固化为 RGB。

### 3. 图片存储
base64 内嵌 JSON，自包含便于对比。大文件场景可后续优化为外部 `media/` 目录引用。

### 4. 空白布局策略
渲染时统一用 `prs.slide_layouts[6]`（Blank）布局，手动放置所有形状，避免原始布局的 placeholder 干扰，保证往返一致性。

### 5. python-pptx 未暴露属性
部分属性（如阴影细节、渐变 stops、自定义几何路径）python-pptx 未直接暴露，通过 `shape._element` 直接操作 lxml 处理。集中在 `utils/xml_helper.py`。

### 6. 对比器的「忽略列表」
往返必然变化的字段（shape_id、自动命名等）加入忽略列表，避免噪声干扰真实缺陷定位。

---

## 验证步骤

1. **环境验证**：`uv run python -c "import pptx; print(pptx.__version__)"` 确认依赖可用
2. **单文件往返**：`uv run python main.py roundtrip input/cover_ending_page.pptx`，检查 `out/json/`、`out/pptx/`、`out/compare/` 产物
3. **全量往返**：`uv run python main.py roundtrip-all`，对 17 个测试文件批量执行
4. **差异分析**：查看 `out/compare/` 报告，定位 parser/renderer 缺陷
5. **迭代修复**：根据差异报告逐个修复解析/渲染逻辑，重跑往返直至差异收敛
6. **测试**：`uv run pytest tests/` 通过

---

## 假设与约定

- **假设**：测试 PPT 中的元素均能被 python-pptx 读取（不涉及 OLE 嵌入对象、SmartArt 等python-pptx 不支持的高级特性，若遇到则记录为「unsupported」并跳过）
- **约定**：所有输出文件 UTF-8 无 BOM 编码
- **约定**：JSON 缩进 2 空格，`ensure_ascii=False`（保留中文）
- **约定**：终端命令统一用 `uv run python main.py ...` 执行
