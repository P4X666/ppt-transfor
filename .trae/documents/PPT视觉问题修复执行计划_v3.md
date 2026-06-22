# PPT 视觉问题修复执行计划 v3

## 摘要

修复用户反馈的三个视觉问题：
1. `content_page_component_table.pptx` 表头应为红色（当前蓝色）—— 根因：表格单元格填充未解析/未渲染
2. `content_page_component_step.pptx` 和 `content_page_component_ladder.pptx` 缺失箭头 —— 根因：`<a:tailEnd>` 箭头线端点未解析/未渲染，且 "Line" 形状被降级为文本框
3. `cover_ending_page.pptx` 缺失图片 —— 根因：图片位于布局/母版上，解析器仅遍历幻灯片级形状

---

## 当前状态分析

### 问题 1：表格表头颜色错误

**原始 XML 结构**（`input/content_page_component_table.pptx` slide1）：
- 表头单元格（第 1 行）：`<a:tcPr><a:solidFill><a:schemeClr val="accent1" /></a:solidFill></a:tcPr>`
- 表体单元格：混合使用 `<a:schemeClr val="accent2"/>"` 和 `<a:srgbClr val="F0F0FA"/>`
- `accent1` 主题色在主题中映射为红色

**当前代码缺陷**：
- [parser/table.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/table.py) `parse_table`：仅解析 text/span_x/span_y，**未解析 `cell.fill`**
- [renderer/table.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/table.py) `render_table`：仅设置 first_row/合并/文本，**未应用 `cell_model.fill`**
- `TableCell` 模型已有 `fill: Optional[Fill] = None` 字段（[schema.py L133](file:///i:/wefor/ppt-transfor/src/ppt_transfor/models/schema.py#L133)），无需改模型
- `add_table` 默认使用蓝色表头样式，由于未回写单元格填充，表头保持默认蓝色

### 问题 2：箭头丢失

**原始 XML 结构**（step.pptx / ladder.pptx 的 "Line" 形状）：
```xml
<a:ln w="25400">
  <a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill>
  <a:miter />
  <a:tailEnd type="triangle" />
</a:ln>
```
- 所有 Line 形状都有 `<a:tailEnd type="triangle" />`（尾部三角箭头），无 `<a:headEnd>`

**当前代码缺陷**：
- `Line` 模型（[schema.py L98-105](file:///i:/wefor/ppt-transfor/src/ppt_transfor/models/schema.py#L98-L105)）仅有 `width/color/dash`，**无箭头字段**
- [parser/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/shape.py) `_parse_line`：仅解析 width/color/dash，**未解析 `<a:headEnd>`/`<a:tailEnd>`**
- [renderer/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/shape.py) `_apply_line`：仅回写 width/color/dash，**未回写箭头 XML**
- "Line" 形状被 `shape.shape_type` 判为 `AUTO_SHAPE`（非 `LINE`），`auto_shape_type=None`
- [renderer/autoshape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/autoshape.py) `auto_shape_type=None` 时返回 None → 降级为文本框，线条和箭头完全丢失

### 问题 3：图片缺失

**原始 XML 结构**（`cover_ending_page.pptx`）：
- slide3/slide4 的 `spTree` **完全为空**（0 个 `<p:sp>`，0 个 `<p:pic>`）
- 图片位于布局上：
  - slide3 → slideLayout3 → `<p:pic>` Picture 1 → image3.png（位置 4724400, 2395361，尺寸 2743200×2067277）
  - slide4 → slideLayout5 → `<p:pic>` Picture 2 → image5.png（位置 1524000, 2821685，尺寸 9144000×1214630）
  - slideMaster1 → `<p:pic>` Image → image1.png（全幅背景，位置 0, 1706，尺寸 12192000×6854588）
- slide0/slide1 也有布局图片（slideLayout1/2 → image2.png），用户反馈"都缺失了图片"

**当前代码缺陷**：
- [parser/slide.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/slide.py) `parse_slide`：仅遍历 `slide.shapes`，**未遍历 `slide.slide_layout.shapes` 和 `slide.slide_layout.slide_master.shapes`**
- 转换后 PPT 使用 Blank 布局，布局/母版继承的图片全部丢失

---

## 修改方案

### 修改 1：表格单元格填充解析与渲染

**文件**：[src/ppt_transfor/parser/table.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/table.py)

**改什么**：在 `parse_table` 中解析每个单元格的 `cell.fill`

**为什么**：当前完全未解析单元格填充，导致表头颜色丢失（默认蓝色）

**怎么改**：
```python
# 在 parse_table 的单元格循环中，cell_model 创建后添加：
from ppt_transfor.parser.shape import _parse_fill
try:
    cell_fill = _parse_fill(cell.fill, cell._tc, prs)
    if cell_fill is not None:
        cell_model.fill = cell_fill
except Exception:
    pass
```
- `cell.fill` 返回 FillFormat，`cell._tc` 是 `<a:tc>` XML 元素
- `_parse_fill` 已支持 SOLID 类型并通过 `parse_color` 固化主题色（accent1 → RGB 红色）
- 传 `cell._tc` 作为 shape_element（noFill 检测对 tcPr 不生效，但 solidFill 正常工作，可接受）

**文件**：[src/ppt_transfor/renderer/table.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/table.py)

**改什么**：在 `render_table` 中应用 `cell_model.fill` 到每个单元格

**为什么**：解析出了填充数据但未回写，表头仍为默认蓝色

**怎么改**：
```python
# 在 render_table 的单元格循环中，文本渲染后添加：
from ppt_transfor.renderer.shape import _apply_fill
if cell_model.fill is not None:
    try:
        _apply_fill(cell.fill, cell_model.fill, None)
    except Exception:
        pass
```
- `cell.fill` 是 python-pptx TableCell 的 FillFormat，支持 `solid()` + `fore_color.rgb`
- `_apply_fill` 已支持 solid/none/background/gradient 类型

### 修改 2：箭头线端点解析与渲染

#### 2a. 扩展 Line 模型

**文件**：[src/ppt_transfor/models/schema.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/models/schema.py#L98-L105)

**改什么**：给 `Line` 模型添加箭头字段

```python
class Line(BaseModel):
    model_config = ConfigDict(extra="allow")
    width: Optional[int] = None
    color: Optional[Color] = None
    dash: Optional[str] = None
    head_arrow_type: Optional[str] = None  # 新增：头部箭头类型（如 "triangle"）
    tail_arrow_type: Optional[str] = None  # 新增：尾部箭头类型（如 "triangle"）
```

#### 2b. 解析箭头线端点

**文件**：[src/ppt_transfor/parser/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/shape.py) `_parse_line` 函数

**改什么**：从 `<a:ln>` XML 中解析 `<a:headEnd>` / `<a:tailEnd>` 的 type 属性

**怎么改**：
```python
def _parse_line(line, prs=None) -> Line | None:
    # ... 现有 width/color/dash 解析逻辑不变 ...

    # 解析箭头线端点
    try:
        from pptx.oxml.ns import qn
        ln_el = line._element.find(qn("a:ln"))
        if ln_el is not None:
            head_end = ln_el.find(qn("a:headEnd"))
            if head_end is not None:
                head_type = head_end.get("type")
                if head_type:
                    model.head_arrow_type = head_type
                    has_value = True
            tail_end = ln_el.find(qn("a:tailEnd"))
            if tail_end is not None:
                tail_type = tail_end.get("type")
                if tail_type:
                    model.tail_arrow_type = tail_type
                    has_value = True
    except Exception:
        pass

    return model if has_value else None
```
- `line._element` 是 spPr 元素，`<a:ln>` 是其子元素
- 仅记录 type 属性（如 "triangle"），忽略 w/len 等次要属性（简化处理）

#### 2c. "Line" auto_shape 重分类为 connector

**文件**：[src/ppt_transfor/parser/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/shape.py) `parse_shape` 函数

**改什么**：当 auto_shape 的 `auto_shape_type=None` 且 line 有箭头时，重分类为 "connector"

**为什么**：这些形状本质是带箭头的直线，当前降级为文本框导致线条和箭头完全丢失

**怎么改**：在 `MSO_SHAPE_TYPE.AUTO_SHAPE` 分支中，`parse_autoshape` 调用后添加判断：
```python
if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
    from ppt_transfor.parser.autoshape import parse_autoshape
    auto_fields = parse_autoshape(shape)

    # 无 auto_shape_type 且有箭头线 → 重分类为 connector
    if auto_fields.get("auto_shape_type") is None:
        line_model = getattr(model, "line", None)
        if line_model is not None and (
            line_model.head_arrow_type or line_model.tail_arrow_type
        ):
            model.shape_type = "connector"
            from ppt_transfor.parser.connector import parse_connector
            conn_fields = parse_connector(shape)
            for k, v in conn_fields.items():
                setattr(model, k, v)
            return model

    model.shape_type = "auto_shape"
    for k, v in auto_fields.items():
        setattr(model, k, v)
    # ... 文本解析不变 ...
    return model
```
- 重分类后 shape_type="connector"，`render_connector` 用 left/top/width/height 计算端点
- `_apply_common_props` 会回写 line 模型（含箭头），箭头通过 2d 的渲染逻辑写入 XML

#### 2d. 渲染箭头线端点

**文件**：[src/ppt_transfor/renderer/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/shape.py) `_apply_line` 函数

**改什么**：回写 `<a:headEnd>` / `<a:tailEnd>` 到 `<a:ln>` XML

**怎么改**：在 `_apply_line` 末尾添加：
```python
# 箭头线端点
if line_model.head_arrow_type or line_model.tail_arrow_type:
    try:
        from lxml import etree
        from pptx.oxml.ns import qn
        ln_el = line._element.find(qn("a:ln"))
        if ln_el is not None:
            # 清除现有箭头
            for tag in ("headEnd", "tailEnd"):
                existing = ln_el.find(qn(f"a:{tag}"))
                if existing is not None:
                    ln_el.remove(existing)
            # 写入新箭头
            if line_model.head_arrow_type:
                head = etree.SubElement(ln_el, qn("a:headEnd"))
                head.set("type", line_model.head_arrow_type)
            if line_model.tail_arrow_type:
                tail = etree.SubElement(ln_el, qn("a:tailEnd"))
                tail.set("type", line_model.tail_arrow_type)
    except Exception:
        pass
```
- 必须在 width/color/dash 设置之后执行（确保 `<a:ln>` 已存在）
- `line._element` 是 spPr，`<a:ln>` 是其子元素

### 修改 3：布局/母版继承图片解析

**文件**：[src/ppt_transfor/parser/slide.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/slide.py)

**改什么**：新增 `_parse_inherited_pictures` 函数，在 `parse_slide` 中调用并前置到 shapes 列表

**为什么**：slide3/slide4 的图片全在布局上，当前完全丢失

**怎么改**：

1. 新增函数：
```python
from pptx.enum.shapes import MSO_SHAPE_TYPE
from ppt_transfor.models.schema import Shape
from ppt_transfor.parser.image import parse_picture

def _parse_inherited_pictures(slide, prs=None) -> list[Shape]:
    """解析布局和母版上的图片形状，返回 Shape 列表。

    转换后 PPT 使用 Blank 布局，布局/母版继承的图片会丢失。
    此函数将布局/母版的 <p:pic> 提取为幻灯片级图片形状，保证视觉保真。
    渲染顺序：母版图片 → 布局图片（前置到 shapes 列表，渲染在后）。
    """
    pictures: list[Shape] = []
    seen_blobs: set[str] = set()

    # 先收集布局图片，再收集母版图片（渲染时母版在后，布局在前）
    # 但前置到 shapes 时需反转顺序：母版图片先渲染（在最底层）
    containers = []
    try:
        containers.append(("layout", slide.slide_layout))
    except Exception:
        pass
    try:
        containers.append(("master", slide.slide_layout.slide_master))
    except Exception:
        pass

    # 母版图片在前（底层），布局图片在后（上层）
    for source, container in reversed(containers):
        try:
            for shape in container.shapes:
                try:
                    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                        continue
                except Exception:
                    continue

                pic_fields = parse_picture(shape)
                if not pic_fields.get("data_base64"):
                    continue

                # 去重：同一图片在母版和布局都有时只保留一个
                blob_key = pic_fields["data_base64"][:100]  # 取前 100 字符作为指纹
                if blob_key in seen_blobs:
                    continue
                seen_blobs.add(blob_key)

                model = Shape(
                    shape_id=str(getattr(shape, "shape_id", "") or ""),
                    name=getattr(shape, "name", None),
                    shape_type="picture",
                    left=int(shape.left) if shape.left is not None else None,
                    top=int(shape.top) if shape.top is not None else None,
                    width=int(shape.width) if shape.width is not None else None,
                    height=int(shape.height) if shape.height is not None else None,
                )
                for k, v in pic_fields.items():
                    setattr(model, k, v)
                pictures.append(model)
        except Exception:
            pass

    return pictures
```

2. 在 `parse_slide` 中调用：
```python
def parse_slide(slide, index: int, prs=None) -> Slide:
    # ... 现有逻辑 ...

    # 遍历形状
    for shape in slide.shapes:
        model.shapes.append(parse_shape(shape, slide, prs))

    # 前置布局/母版继承的图片（渲染在底层，保证视觉保真）
    inherited_pictures = _parse_inherited_pictures(slide, prs)
    if inherited_pictures:
        model.shapes = inherited_pictures + model.shapes

    # 后处理：修正低对比度文本
    _ensure_text_visibility(model)

    return model
```

**设计决策**：
- 母版图片在底层（先渲染），布局图片在上层（后渲染），符合 PowerPoint 渲染顺序
- 用 base64 前 100 字符去重，避免同一图片在母版和布局重复添加
- 仅提取 PICTURE 类型，不提取 placeholder/sp 等其他布局形状
- 前置到 shapes 列表，确保渲染在 slide 级形状之前（底层）

---

## 假设与决策

1. **表格单元格 noFill 检测**：`_parse_fill` 的 noFill 检查针对 `spPr`，对 `tcPr` 不生效。可接受 —— 无填充的单元格 `fill.type` 返回 None，自然不记录 fill 字段，行为正确。

2. **箭头仅记录 type 属性**：忽略 `w`（宽度）和 `len`（长度）属性，仅保留 `type`（如 "triangle"）。视觉上箭头类型是最关键信息，宽度和长度使用默认值可接受。

3. **"Line" auto_shape 重分类条件**：仅当 `auto_shape_type=None` 且 line 有箭头时重分类为 connector。无箭头的无类型 auto_shape 保持现有降级为文本框的行为（不在本次修复范围）。

4. **继承图片添加到所有幻灯片**：原始 PPT 中布局/母版图片确实出现在所有使用该布局/母版的幻灯片上，因此全部添加是正确的。转换后 PPT 使用 Blank 布局，不添加则图片丢失。

5. **继承图片去重**：用 base64 前 100 字符作为指纹，避免同一图片在母版和布局重复添加。不同图片不会误去重。

6. **comparator 无需修改**：三个修复都使原始 JSON 和转换后 JSON 结构一致，现有等价规则已覆盖。表格 fill、箭头字段、继承图片在往返后结构相同，不产生新差异。

---

## 验证步骤

### 1. 单文件视觉验证

```bash
# 重新生成 JSON 和 PPT
uv run python main.py convert content_page_component_table
uv run python main.py convert content_page_component_step
uv run python main.py convert content_page_component_ladder
uv run python main.py convert cover_ending_page
```

打开 `out/pptx/` 下对应文件，视觉对比：
- `content_page_component_table.pptx`：表头应为红色（accent1 主题色）
- `content_page_component_step.pptx`：所有 Line 形状尾部应有三角箭头
- `content_page_component_ladder.pptx`：Line 形状尾部应有三角箭头
- `cover_ending_page.pptx`：4 页都应有图片（母版背景图 + 布局图片）

### 2. 往返测试

```bash
# 全部 17 个 PPT 往返测试，期望 0 差异
uv run python main.py roundtrip-all
```

```bash
# 单元测试
uv run pytest tests/test_roundtrip.py -v
```

### 3. JSON 结构验证

检查 `out/json/` 下对应 JSON：
- `content_page_component_table.json`：`table.cells[0][*].fill` 应有 `{"type":"solid","color":{"type":"rgb","value":"..."}}`
- `content_page_component_step.json`：Line 形状的 `shape_type` 应为 `"connector"`，`line.tail_arrow_type` 应为 `"triangle"`
- `content_page_component_ladder.json`：同上
- `cover_ending_page.json`：slides[2] 和 slides[3] 的 `shapes` 不再为空，应包含 picture 类型形状
