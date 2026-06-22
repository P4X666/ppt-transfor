# PPT 视觉问题修复执行计划 v2

## 一、概要

修复 input/ 下 PPTX 在解析→渲染后出现的三类视觉差异，并全量扫描所有 input PPT 避免同类问题遗漏。

目标文件与现象：

1. `content_page_component_multicolumn.pptx`：转换后背景呈蓝色，与原始 PPT 差异大。
2. `content_page_component_multi_ring.pptx` / `content_page_component_radial_list.pptx`：原始彩色圆环（Chart）转换后只剩圆形/空白。
3. `section_page.pptx`：文本内容超出屏幕，缺少自动换行。

## 二、当前状态分析（基于代码审查）

### 2.1 已落地修复

- **渐变填充（形状级）**：
  - 模型：`src/ppt_transfor/models/schema.py` 已扩展 `Fill.gradient_type / gradient_angle / gradient_stops`。
  - 解析：`src/ppt_transfor/parser/shape.py` 已能读取 `<a:gradFill>` 的 stops、`<a:lin>` 角度、`<a:path>` 类型。
  - 渲染：`src/ppt_transfor/renderer/shape.py` 的 `_apply_gradient_fill` 已通过 XML 直接回写 `<a:gsLst>` 与 `<a:lin>`。

- **文本自动换行**：
  - `src/ppt_transfor/renderer/text.py` 的 `render_text_frame` 已支持 `default_word_wrap` 参数。
  - `src/ppt_transfor/renderer/shape.py` 对 `text_box / placeholder / auto_shape / 未知类型` 已传入 `default_word_wrap=True`。

- **Chart 保留框架**：
  - 模型：`Shape.chart_xml / chart_part` 已存在。
  - 解析：`src/ppt_transfor/parser/shape.py` 已对 `MSO_SHAPE_TYPE.CHART` 记录 `chart_xml` 与 `chart_part`。
  - 渲染：`src/ppt_transfor/renderer/shape.py` 对 `shape_type == "chart"` 直接返回 `None`，不生成占位文本框。
  - 新增 `src/ppt_transfor/renderer/chart_preserve.py`，在 `render_presentation` 保存后通过 zip 级后处理插入原始 `<p:graphicFrame>`、复制 chart part 与依赖、更新 slide rels 和 `[Content_Types].xml`。
  - `main.py` 的 `roundtrip` / `roundtrip-all` 已传入 `source_pptx_path`。

### 2.2 仍待修复/验证

- **幻灯片背景渐变缺失（multicolumn 蓝色背景根因）**：
  - `src/ppt_transfor/utils/inheritance.py` 的 `resolve_background` 只返回背景类型与单色，不解析渐变 stops/角度。
  - `src/ppt_transfor/renderer/slide.py` 的 `_apply_background` 只处理 `solid` 背景，遇到 `gradient` 背景不设置，导致回退到 Blank 布局的默认背景（可能呈现蓝色）。
  - 若 multicolumn 的背景是幻灯片级渐变，则必须扩展 `Background` 模型并支持渐变渲染。

- **Chart 保留后解析可能抛出 `KeyError: 'rId7'`**：
  - 原因推断：chart 后处理写入的 slide rels / chart rels 与 python-pptx 后续解析预期不一致，或 chart XML 中除 `<c:chart r:id>` 外还有其他 `r:id` 引用未更新。
  - 需要运行 roundtrip 复现并加固 relationship 处理逻辑。

- **Chart 依赖 part 的 content type 可能遗漏**：
  - 当前 `_content_type_for_part` 只覆盖 xlsx/图片等，对 chart style (`style.xml`) 和 color (`colors.xml`) 的 content type 未显式登记。
  - 虽然 chart 通常只要求 chart XML 本身能打开，但为保险起见需要为 `style.xml` 和 `colors.xml` 添加 Override。

- **测试未验证 chart 保留**：
  - `tests/test_roundtrip.py` 调用 `render_presentation` 时未传入 `source_pptx_path`，导致 chart 保留在 pytest 中不生效，无法回归防护。

- **全量扫描文本换行**：
  - 当前 `default_word_wrap=True` 已覆盖主要形状，但组合形状内的文本框是否继承该兜底需要验证。

## 三、修改方案

### 3.1 幻灯片背景支持渐变填充

**文件**：`src/ppt_transfor/models/schema.py`

- 扩展 `Background` 模型，复用 `Fill` 的渐变字段：

```python
class Background(BaseModel):
    type: str = "none"
    color: Optional[Color] = None
    # 渐变背景支持
    gradient_type: Optional[str] = None
    gradient_angle: Optional[float] = None
    gradient_stops: list[GradientStop] = []
```

**文件**：`src/ppt_transfor/utils/inheritance.py`

- 在 `resolve_background` 中，当背景类型为 `GRADIENT` 时，调用 `_parse_gradient_from_xml`（或内联等价逻辑）从 `<p:bg>/<p:bgPr>/<a:gradFill>` 提取 stops/角度/类型，写入 `Background.gradient_*`。
- 兼容 slide / layout / master 三层继承链，取第一个非 BACKGROUND 的渐变或纯色。

**文件**：`src/ppt_transfor/renderer/slide.py`

- 扩展 `_apply_background`：当 `bg.type == "gradient"` 时，调用 slide background 的 `fill.gradient()` 创建骨架，然后复用/改造 `_apply_gradient_fill` 逻辑（或提取公共函数到 `utils/xml_helper.py`）向 `<a:bgPr>/<a:gradFill>` 写入 stops 与方向。
- 保持当前 `solid` 与 `none` 行为不变。

### 3.2 Chart 保留鲁棒性加固

**文件**：`src/ppt_transfor/renderer/chart_preserve.py`

- 修复/验证 relationship 处理：
  1. `_process_slide_charts` 中确保 `slide_rels_path` 即使原本不存在也能生成新的 relationships（当前已读 `zout` 现有 rels，逻辑正确，需运行验证）。
  2. 在更新 `<c:chart r:id="...">` 后，扫描 `graphic_frame` 中所有 `{NS_R}id` 属性，确认没有其他未解析的 rId 指向缺失 part；若有，记录 warning 并保留原 rId（或删除该属性），避免 KeyError。
  3. 写入新的 slide rels 前，先校验所有 slide XML 中引用的 rId 都在 rels 中存在；若缺失，从原 rels 补齐。
- 为 chart style / colors part 补 content type：
  - 在 `_content_type_for_part` 中增加 `.xml` 的兜底：若 rel_type 为 `http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartStyle`，content type 为 `application/vnd.openxmlformats-officedocument.chartstyle+xml`；若为 `.../chartColorStyle`，content type 为 `application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml`（按 OpenXML 实际类型填写）。
- 在 `_copy_chart_dependencies` 中复制完依赖后，记录 chart rels 中所有 relationship 的 `TargetMode="External"` 情况，外部链接不复制，仅保留 Target。

### 3.3 文本换行兜底验证

**文件**：`src/ppt_transfor/renderer/shape.py`

- 当前已覆盖 `text_box / placeholder / auto_shape / unknown`。计划执行时通过 `roundtrip-all` 验证 `section_page.pptx` 是否仍有文本溢出。
- 若组合形状（`group`）内的子文本框仍溢出，检查 `src/ppt_transfor/renderer/group.py`，确保渲染子形状时调用 `render_shape` 的 `default_word_wrap=True` 逻辑生效（`render_shape` 内部已处理，通常无需额外改动）。

### 3.4 测试补全

**文件**：`tests/test_roundtrip.py`

- 修改 `render_presentation(original_model, converted_pptx)` 为：

```python
render_presentation(original_model, converted_pptx, source_pptx_path=pptx_path)
```

- 这样 pytest 路径也能触发 chart 保留，避免未来回归。

## 四、执行顺序

1. **先验证**：运行 `uv run pytest` 与 `uv run python main.py roundtrip-all`，收集当前失败文件与错误信息。
2. **修复幻灯片背景渐变**：完成 3.1。
3. **修复 Chart 保留问题**：完成 3.2，并针对 `multi_ring` / `radial_list` 单文件反复 roundtrip 直到能正常解析。
4. **验证文本换行**：完成 3.3，确认 `section_page.pptx` 无溢出。
5. **更新测试**：完成 3.4。
6. **全量回归**：再次运行 `uv run pytest` 与 `uv run python main.py roundtrip-all`。
7. **可选视觉对比**：若环境有 LibreOffice，运行 `uv run python main.py roundtrip-all --visual` 抽查。

## 五、验收标准

1. `content_page_component_multicolumn.pptx` roundtrip 后背景与原始文件一致（无蓝色异常）。
2. `content_page_component_multi_ring.pptx` / `content_page_component_radial_list.pptx` roundtrip 后保留彩色圆环，转换后 PPTX 可被 python-pptx 正常解析，无 `KeyError`。
3. `section_page.pptx` roundtrip 后文本自动换行，不超出屏幕。
4. `uv run pytest` 全部通过。
5. `uv run python main.py roundtrip-all` 无新增严重差异；渐变 stops / chart_xml 等结构性差异经确认属于正确修复。

## 六、风险与回退

1. **幻灯片渐变背景 XML 路径差异**：部分 PPT 的渐变定义在 `<p:bgRef>` 而非 `<p:bgPr>`，若遇到则按 reference 处理：先尝试复用 shape 渐变逻辑，无法解析时记录 warning 并保持当前行为。
2. **Chart 保留破坏其他 slide rels**：若加固后仍出现 KeyError，可临时对特定文件跳过 chart 保留，或回退到“chart 降级为图片”方案（需引入截图，工作量大，不作为首选）。
3. **默认换行改变特殊形状**：若验收中发现某 auto_shape 长文本原意不换行，可针对该 shape 名/类型加白名单，但当前需求优先保证可见性。
