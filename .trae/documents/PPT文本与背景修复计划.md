# PPT 文本与背景修复计划

## 一、摘要

修复 `conent_page_component_card.pptx` 与 `content_page_component_calendar.pptx` 在往返转换后出现的文本未居中、标题未换行、背景色错误（大面积变黑）等问题；并通过全量 roundtrip + 视觉对比扫描其余 15 个输入文件，定位并修复其他严重视觉不一致。

## 二、当前状态分析

基于对以下关键文件的阅读：
- [src/ppt_transfor/parser/text.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/text.py)
- [src/ppt_transfor/renderer/text.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/text.py)
- [src/ppt_transfor/parser/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/shape.py)
- [src/ppt_transfor/renderer/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/shape.py)
- [src/ppt_transfor/utils/inheritance.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/utils/inheritance.py)
- [src/ppt_transfor/parser/slide.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/slide.py)
- [src/ppt_transfor/renderer/slide.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/slide.py)
- [src/ppt_transfor/utils/color.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/utils/color.py)
- [src/ppt_transfor/parser/table.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/table.py)
- [main.py](file:///i:/wefor/ppt-transfor/main.py)
- [tests/test_roundtrip.py](file:///i:/wefor/ppt-transfor/tests/test_roundtrip.py)

已具备的能力：
1. `BACKGROUND` 填充类型已被保留并调用 `fill.background()` 回写。
2. 文本对齐从段落 XML `<a:pPr algn>` 和 shape 级 `<a:lstStyle>/<a:lvl1pPr>` 提取并合并到段落模型。
3. `word_wrap`、`auto_size`、`vertical_anchor` 已解析并渲染。
4. 主题色/Scheme 色可通过 theme part 的 `clrScheme` 固化为 RGB。
5. `cap`（全大写）已解析并渲染。

仍可能存在的缺陷：
1. **`<a:bodyPr>` 解析不完整**：`wrap` 属性、内部边距 `lIns/tIns/rIns/bIns` 未解析/渲染。若 python-pptx 的 `tf.word_wrap` 未正确暴露原始 wrap 设置，会导致标题不换行。
2. **文本对齐兜底不足**：`auto_shape` 降级为 `text_box` 时，add_textbox 默认段落为左对齐；若模型对齐为 `None`（例如继承自 theme/layout 且未被固化），文本会偏左。
3. **BACKGROUND 填充的视觉陷阱**：`fill.background()` 依赖当前 slide/layout 背景。若 Blank 布局自带深色背景，或 slide 背景未被正确设置，BACKGROUND 形状会显示为黑色。
4. **SOLID 填充颜色丢失后的默认黑色**：`_parse_fill` 中 SOLID 类型若 `parse_color` 返回 `None`（如 scheme 色解析失败），模型 color 为 `None`；渲染时只调用 `fill.solid()` 不设置颜色，python-pptx 可能使用默认黑色填充。
5. **表格单元格未继承默认样式**：`parser/table.py` 直接调用 `parse_text_frame(cell.text_frame, prs)`，未传入单元格级别的默认文本属性，导致表格内文本对齐/字号可能丢失。
6. **auto_shape 降级居中对齐兜底缺失**：[renderer/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/shape.py#L163-L172) 降级为 textbox 后未显式应用 shape 的默认居中对齐。

## 三、拟议修改

### 3.1 增强 `<a:bodyPr>` 解析（影响文本换行与垂直对齐）

**文件**：`src/ppt_transfor/parser/text.py`

- 新增 `_parse_bodyPr_wrap(tf_element)`：直接读取 `<a:bodyPr wrap="square|none">`，当 `tf.word_wrap` 为 `None` 或不可靠时作为补充。
- 新增 `_parse_bodyPr_insets(tf_element)`：读取 `lIns/tIns/rIns/bIns`（单位 EMU），存入 `Text.margin_left/top/right/bottom`。
- 在 `parse_text_frame` 中调用上述函数。

**文件**：`src/ppt_transfor/models/schema.py`

- 在 `Text` 模型新增可选字段：`margin_left`、`margin_top`、`margin_right`、`margin_bottom`（EMU 整数）。

**文件**：`src/ppt_transfor/renderer/text.py`

- 新增 `_apply_insets(tf, text_model)`：将 margin 回写到 `<a:bodyPr>` 的 `lIns/tIns/rIns/bIns`（通过 XML 操作，因为 python-pptx 未暴露 API）。
- 在 `render_text_frame` 中调用。

### 3.2 文本对齐更强兜底

**文件**：`src/ppt_transfor/parser/text.py`

- 在 `parse_text_frame` 的 `merged_props` 合并逻辑后，确保 shape 默认对齐（来自 `extract_txbody_default_props`）能够覆盖 `inherited_props` 中的 `None` 值。当前逻辑已实现，但需验证 `extract_txbody_default_props` 对非 placeholder 文本框也能正确返回对齐。

**文件**：`src/ppt_transfor/renderer/shape.py`

- 在 `auto_shape` 降级为 `text_box` 的分支中，创建 textbox 并渲染文本后，若模型 `text` 中存在段落对齐，显式应用第一段对齐（add_textbox 默认创建空段落，其默认对齐为 LEFT）。
- 更通用的兜底：在 `_apply_common_props` 之后，若 shape 含文本且模型 `text` 的默认对齐不为 `None`，遍历并应用所有段落对齐。

### 3.3 修复 BACKGROUND/SOLID 填充导致的大面积黑色

**文件**：`src/ppt_transfor/utils/color.py` 与 `src/ppt_transfor/utils/inheritance.py`

- 增强 `_get_theme_element` 的健壮性：若 `master.part.rels` 找不到 theme，尝试 `presentation.part.rels` 兜底。
- 在 `parse_color` 中，SCHEME 类型解析失败时，至少返回 `Color(type="theme", value=...)`，避免颜色完全丢失。

**文件**：`src/ppt_transfor/parser/shape.py`

- `_parse_fill` 中 SOLID 类型且 `color` 为 `None` 时，将 `Fill.type` 降级为 `"none"` 或标记 `"solid_no_color"`，避免渲染时调用 `fill.solid()` 并使用默认黑色。推荐做法：若无法解析颜色，返回 `Fill(type="none")`。

**文件**：`src/ppt_transfor/renderer/shape.py`

- `_apply_fill` 中 `fill_model.type == "solid"` 且 `fill_model.color is None` 时，跳过 `fill.solid()`，不修改填充，防止默认黑色。
- `fill_model.type == "background"` 时，优先尝试将已解析的 slide 背景色作为 SOLID 填充应用（若 slide 背景为 solid 且 color 已知）；否则再调用 `fill.background()`。

**文件**：`src/ppt_transfor/renderer/slide.py`

- `_apply_background` 中，若背景模型为 `solid` 但 `color` 为 `None`，跳过设置，避免默认黑色背景。

### 3.4 表格单元格默认文本样式继承

**文件**：`src/ppt_transfor/parser/table.py`

- 对每个单元格调用 `parse_text_frame` 前，先使用 `extract_txbody_default_props(cell.text_frame._element, prs)` 提取单元格默认文本属性，并作为 `inherited_props` 传入。

### 3.5 全量扫描其他不一致

**文件**：`main.py`（仅使用现有命令，不修改代码）

- 运行 `uv run python main.py roundtrip-all` 生成 JSON diff 报告。
- 若已安装 LibreOffice，运行 `uv run python main.py roundtrip-all --visual` 生成视觉差异。
- 对视觉差异比例 > 5% 或 JSON diff 数量显著（> 10）的文件，使用 `inspect_*.py` 脚本或新增临时检查脚本定位根因。

## 四、假设与决策

1. **判定标准**：将“严重不一致”定义为视觉差异比例 > 5% 或 JSON diff 中关键字段（`fill.color`、`text.alignment`、`shape.left/top/width/height`）存在明显差异。
2. **主题色处理**：继续沿用“主题色固化为 RGB”的策略，以视觉保真为优先；若某主题色无法解析，则降级为不设置颜色，而不是默认黑色。
3. **BLANK 布局**：渲染统一使用 Blank 布局的决策不变；若 Blank 布局背景导致 BACKGROUND 填充异常，将通过显式设置 slide 背景来解决。
4. **中间产物**：保留 `out/json/`、`out/pptx/`、`out/compare/`、`out/visual/` 中间产物，便于人工核对。
5. **运行环境**：假设用户环境已安装 `uv`；视觉对比依赖 LibreOffice，若未安装则仅使用 JSON diff。

## 五、验证步骤

1. **单元验证**：针对 `conent_page_component_card.pptx` 和 `content_page_component_calendar.pptx` 分别运行：
   ```bash
   uv run python main.py roundtrip input/conent_page_component_card.pptx --visual
   uv run python main.py roundtrip input/content_page_component_calendar.pptx --visual
   ```
   期望：JSON diff 数量为 0 或仅包含可忽略字段；视觉差异比例 < 1%。

2. **全量回归**：运行 pytest：
   ```bash
   uv run pytest tests/test_roundtrip.py -v
   ```
   期望：所有 17 个文件通过（diff_count == 0）。

3. **全量视觉扫描**（若 LibreOffice 可用）：
   ```bash
   uv run python main.py roundtrip-all --visual
   ```
   期望：无文件平均视觉差异 > 5%；对 > 1% 的文件生成热图并人工复核。

4. **手动抽查**：使用 PowerPoint 打开转换后的 `content_page_component_calendar.pptx` 与 `conent_page_component_card.pptx`，确认：
   - 卡片标题文本居中且按原始宽度换行。
   - 日历仅顶部条和月份标签为黑色背景，其余格子为浅色背景。
