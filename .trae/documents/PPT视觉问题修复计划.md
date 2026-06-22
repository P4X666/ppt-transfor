# PPT 视觉问题修复执行计划

## 一、概要

修复 input/ 下 PPTX 在解析→渲染后出现的三类视觉差异：

1. `content_page_component_multicolumn.pptx`：转换后背景呈蓝色（渐变填充 stops/角度丢失）。
2. `content_page_component_multi_ring.pptx` / `content_page_component_radial_list.pptx`：转换后只能看到圆形/背景，丢失彩色圆环（实际是 Chart 图形被降级为空白文本框）。
3. `section_page.pptx`：转换后文本内容超出屏幕（文本框未自动换行）。

同时扫描所有 input PPT，对同类问题统一修复。

## 二、当前状态分析

### 2.1 代码结构（已探索确认）

- 解析入口：`src/ppt_transfor/parser/presentation.py` → `parse_slide` → `parse_shape`
- 渲染入口：`src/ppt_transfor/renderer/presentation.py` → `render_slide` → `render_shape`
- 数据模型：`src/ppt_transfor/models/schema.py`
- 填充解析：`src/ppt_transfor/parser/shape.py` 中 `_parse_fill` / `_parse_gradient_from_xml`
- 填充渲染：`src/ppt_transfor/renderer/shape.py` 中 `_apply_fill` / `_apply_gradient_fill`
- 文本渲染：`src/ppt_transfor/renderer/text.py` 中 `render_text_frame`
- 对比器：`src/ppt_transfor/comparator/differ.py`（已将 chart/placeholder/auto_shape 与 text_box 视为等价）

### 2.2 已落地修复

- **渐变填充**：模型已扩展 `Fill.gradient_type / gradient_angle / gradient_stops`，解析端已读取 `<a:gradFill>` 的 stops/`<a:lin>`，渲染端已通过 XML 直接回写 `<a:gsLst>` 与 `<a:lin>`。
- **文本换行**：`render_text_frame` 已支持 `default_word_wrap` 参数，文本框/占位符渲染时已传入 `True`。

### 2.3 仍待完成

- **Chart 图形保留**：`Shape.chart_xml` 字段已在模型中，但 `parse_shape` 未对 `MSO_SHAPE_TYPE.CHART` 赋值；`render_shape` 对 chart 仍降级为 `add_textbox`；最终 PPTX 中缺少 chart part 与 `<p:graphicFrame>`。
- **全量扫描验证**：需要跑通全部 input 的 roundtrip，确认没有新增回归。

## 三、修改方案

### 3.1 扩展数据模型（chart 保留需要）

**文件**：`src/ppt_transfor/models/schema.py`

在 `Shape` 模型新增字段：

```python
# 图表特有：原始 chart part 路径，例如 "ppt/charts/chart1.xml"
chart_part: Optional[str] = None
```

`chart_xml` 已存在，用于保存原始 `<p:graphicFrame>` 的 XML 字符串。新增 `chart_part` 是为了在渲染阶段快速定位需要复制到目标 PPTX 的 chart part。

### 3.2 解析端识别并记录 Chart

**文件**：`src/ppt_transfor/parser/shape.py`

在 `parse_shape` 的类型分发中，显式处理 `MSO_SHAPE_TYPE.CHART`：

1. 设置 `model.shape_type = "chart"`。
2. 将 `<p:graphicFrame>` 元素序列化为字符串，保存到 `model.chart_xml`。
3. 从所属 slide part 的 relationships 中找到该 graphicFrame 引用的 chart relationship（`<a:graphicData><c:chart r:id="..."/></a:graphicData>` 中的 `r:id`），记录其 target part 路径（去掉开头的 `/`）到 `model.chart_part`。

若无法读取 relationship，则仅保存 `chart_xml`，渲染阶段尝试通过 `r:id` 在原始 PPTX 中反查；若仍失败则跳过 chart，保持当前降级为文本框的行为。

### 3.3 渲染端跳过 Chart 占位符

**文件**：`src/ppt_transfor/renderer/shape.py`

在 `render_shape` 中新增 `chart` 分支：

```python
elif model.shape_type == "chart":
    # chart 通过保存后的 zip 后处理插入，这里不生成占位文本框
    return None
```

`render_slide` 对 `None` 返回值不做处理，继续渲染下一个形状。

### 3.4 Chart XML 与 Part 级保留（核心）

**新增文件**：`src/ppt_transfor/renderer/chart_preserve.py`

实现 `apply_chart_preservation(output_pptx_path, source_pptx_path, model)`，在 `render_presentation` 保存文件后对目标 PPTX 进行 zip 级后处理：

1. **准备**
   - 以 `zipfile.ZipFile(..., mode='a')` 打开已保存的目标 PPTX。
   - 以只读方式打开原始 PPTX zip。
   - 遍历模型，收集所有 `shape_type == "chart"` 且 `chart_xml` 非空的 shape，按 slide 索引分组。

2. **插入 `<p:graphicFrame>` 到 slide XML**
   - 读取 `ppt/slides/slide{N}.xml`。
   - 找到 `<p:spTree>`，收集其中形状子元素（`p:sp`、`p:pic`、`p:graphicFrame`、`p:grpSp`、`p:cxnSp`）。
   - 遍历模型 `slide.shapes`：
     - 遇到非 chart 形状时，维护 `shape_idx`（指向当前对应的目标 slide 中形状元素）。
     - 遇到 chart 形状时，将 `chart_xml` 解析为 lxml 元素，替换其内部 `<c:chart r:id="..."/>` 的 `r:id` 为新分配 slide-chart relationship 的 rId；然后在该 slide 的 `shape_elements[shape_idx]` 之前插入（若 `shape_idx == len(shape_elements)` 则 append）。
   - 连续多个 chart 时，`addprevious` 会依次排在前面，天然保持原始顺序。

3. **复制 chart part 及其依赖**
   - 对每个 chart：
     - 在目标 PPTX 中分配新的唯一 part 名，如 `ppt/charts/chart{next_id}.xml`。
     - 从原始 PPTX 复制 chart XML 内容到该 part。
     - 读取原始 chart 的 rels 文件（`ppt/charts/_rels/chart{orig_id}.xml.rels`），复制其中 relationship 指向的 part（当前主要是 `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`）到目标 PPTX，并重命名为唯一名称。
     - 更新目标 chart rels 文件中的 `Target` 指向重命名后的 embedding。
     - 在 slide rels 文件（`ppt/slides/_rels/slide{N}.xml.rels`）中新增一条 `chart` relationship，指向新的 chart part，rId 唯一。
     - 将 slide rels 中新的 rId 回写到刚插入的 `<c:chart r:id="..."/>`。

4. **更新 `[Content_Types].xml`**
   - 为新增的 chart part 添加 `application/vnd.openxmlformats-officedocument.chartml.chart+xml` Override。
   - 为新增的 embedding xlsx 添加 `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` Override。

5. **兜底与错误处理**
   - 任何 chart 处理失败时，记录 warning 并跳过该 chart，不影响其他形状。
   - 若原始 PPTX 找不到或 `chart_xml` 为空，直接返回，不修改输出文件。

### 3.5 渲染入口调用 Chart 保留

**文件**：`src/ppt_transfor/renderer/presentation.py`

1. 修改 `render_presentation` 签名，新增可选参数 `source_pptx_path: str | Path | None = None`。
2. 当 `source_pptx_path` 为 `None` 时，尝试通过 `Path("input") / model.source_file` 定位原始文件；存在则使用。
3. 在 `prs.save(...)` 之后，若原始文件存在，调用：
   ```python
   from ppt_transfor.renderer.chart_preserve import apply_chart_preservation
   apply_chart_preservation(output_path, source_pptx_path, model)
   ```

**文件**：`main.py`

在 `roundtrip` / `roundtrip-all` 调用 `render_presentation` 时，传入原始 PPTX 路径，确保 chart 保留生效。

### 3.6 文本换行兜底扩展到所有带文本形状

**文件**：`src/ppt_transfor/renderer/shape.py`

当前只有 `text_box` / `placeholder` 传入 `default_word_wrap=True`。为覆盖所有可能出现的长文本溢出：

- `auto_shape` 渲染文本时，调用 `render_text_frame(..., default_alignment="CENTER", default_word_wrap=True)`。
- 未知类型兜底分支也传入 `default_word_wrap=True`。

`text_model.word_wrap` 显式值仍优先，不会破坏原始“不换行”的设置。

## 四、验收标准

1. `content_page_component_multicolumn.pptx` roundtrip 后背景为灰色渐变，JSON diff 不出现蓝色/默认渐变相关差异。
2. `content_page_component_multi_ring.pptx` / `content_page_component_radial_list.pptx` roundtrip 后输出 PPTX 中能看到彩色圆环，PowerPoint / python-pptx 可正常打开。
3. `section_page.pptx` roundtrip 后文本正常换行、不超出屏幕。
4. `uv run python main.py roundtrip-all` 无新增严重差异；对新增 `chart_xml` / `chart_part` 字段以及渐变 stops 的结构性 diff，确认属于正确修复。
5. 运行 `uv run python tests/detect_visibility_issues.py`（如存在）无新增黑底黑字/白底白字问题。

## 五、验证步骤

1. 单元验证：
   ```bash
   uv run python main.py roundtrip input/content_page_component_multicolumn.pptx
   uv run python main.py roundtrip input/content_page_component_multi_ring.pptx
   uv run python main.py roundtrip input/content_page_component_radial_list.pptx
   uv run python python main.py roundtrip input/section_page.pptx
   ```
2. 批量验证：
   ```bash
   uv run python main.py roundtrip-all
   ```
3. 可选视觉验证：
   ```bash
   uv run python main.py roundtrip-all --visual
   ```
4. 人工抽查输出文件 `out/pptx/*.pptx` 中 multi_ring / radial_list / section_page / multicolumn 的视觉效果。

## 六、风险与回退

1. **Chart XML 插入位置错误**：若 spTree 中前缀子元素导致索引计算错误，可能破坏 slide XML。回退：在 `apply_chart_preservation` 中捕获异常，删除错误写入并跳过该 chart。
2. **rId / part 名冲突**：分配新 rId 与 part 名时需扫描现有文件取最大值。若冲突，PowerPoint 可能打不开。回退：生成唯一名时使用 `max_existing + 1` 并校验不存在。
3. **Chart 依赖 part 未复制全**：当前 chart rels 只有 embedding xlsx；实现时会通用遍历 chart rels 复制所有 target part。若出现未预见的 reltype，记录 warning 并跳过。
4. **默认换行影响特殊形状**：若某 auto_shape 原意不换行且模型 `word_wrap` 为 `None`，设置 `True` 会改变行为。概率低，且符合 PowerPoint 默认；若验收时发现 regression，可改为仅对 text_box / placeholder 开启。
