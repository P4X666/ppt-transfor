# PPT 转换视觉差异修复计划

## 概述

修复 PPT 转换服务中"JSON 往返 0 差异但视觉差异巨大"的问题。核心根因：解析器丢弃了所有继承自 layout/master/theme 的属性（背景、对齐、颜色、字号），对比器又过度等价化掩盖了真实差异。

**用户确认的方案**：
1. 仅修 JSON 解析 + 对比器（不引入图像对比层）
2. 主题色固化为 RGB（牺牲主题色语义换 100% 视觉保真）

---

## 当前状态分析

### 问题根因（基于代码探索）

| 问题 | 根因 | 关键代码位置 |
|------|------|-------------|
| A 黑色背景丢失 | 解析不读继承（slide→layout→master），渲染用空白模板 | [parser/slide.py:19-20,32-34](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/slide.py) |
| B title 居中丢失 | alignment 继承为 None 被丢弃，text_box 默认左对齐 | [parser/text.py:13-17,71](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/text.py) |
| C 浅灰色辅色丢失 | 继承颜色 None 丢弃 + 主题色索引映射错 | [utils/color.py:38-39](file:///i:/wefor/ppt-transfor/src/ppt_transfor/utils/color.py) |
| D calendar 溢出 | group chOff/chExt 未捕获 + 字号缺失撑大 SHAPE_TO_FIT_TEXT | [renderer/group.py:36-44](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/group.py) |
| E 对比器过度等价 | 6 条等价规则掩盖真实差异 | [comparator/differ.py:15-20,82-84,89-90,102-105](file:///i:/wefor/ppt-transfor/src/ppt_transfor/comparator/differ.py) |

### 关键证据
- `out/json/cover_ending_page.json` 中 layout_name 为 "Logo Intro A: Black"，但 background 记为 `{"type": "none"}`
- 所有 run 的 font 都是 `{}`（空），字号/颜色/字体名全部缺失
- 对比器把 `layout_name` 列入 IGNORED_KEYS，把 None≡CENTER、placeholder≡text_box 等等价化

---

## 实施步骤

### 步骤 1：扩展数据模型
**文件**：[src/ppt_transfor/models/schema.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/models/schema.py)

- `Shape` 增加 `child_offset` 和 `child_extent` 字段（group 的子坐标系 chOff/chExt）
  ```python
  # 组合特有：子坐标系（chOff/chExt）
  child_offset: Optional[tuple[int, int]] = None  # (x, y) EMU
  child_extent: Optional[tuple[int, int]] = None  # (cx, cy) EMU
  ```
  > 注：pydantic tuple 序列化为 JSON 数组，反序列化自动还原

### 步骤 2：实现继承解析工具
**文件**：[src/ppt_transfor/utils/inheritance.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/utils/inheritance.py)（新建）

这是本次修复的核心工具，提供三个继承解析能力：

```python
def resolve_background(slide) -> Background:
    """解析幻灯片背景，沿继承链向上查找。

    顺序：slide.background → slide.slide_layout.background → slide.slide_layout.slide_master.background
    取第一个 SOLID 填充作为真实背景。
    """

def resolve_placeholder_props(shape) -> dict:
    """解析 placeholder 形状的继承属性（对齐、字体、字号、颜色）。

    通过 shape.placeholder_format.idx 匹配 layout/master 的 placeholder，
    读取其段落对齐、run 字体属性。
    返回 dict: { "alignment": str|None, "font_size": int|None, "font_color": Color|None, "font_name": str|None }
    """

def resolve_theme_color(color_format, presentation) -> Color:
    """将主题色解析为具体 RGB。

    通过 presentation.slide_masters[0].element 获取 theme part，
    根据 schemeClr 名称（如 accent1/tx1）查 clrScheme，固化为 type="rgb"。
    """
```

### 步骤 3：修复背景解析
**文件**：[src/ppt_transfor/parser/slide.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/slide.py)

- `_parse_background` 改为调用 `resolve_background`，沿继承链解析
- 当 slide 级别 fill.type 为 None 或 BACKGROUND 时，向上查 layout 和 master
- 取到 SOLID 填充后记录真实颜色

### 步骤 4：修复颜色解析（主题色固化为 RGB）
**文件**：[src/ppt_transfor/utils/color.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/utils/color.py)

- `parse_color` 增加参数 `presentation`，用于解析主题色
- 当 color_type 为 SCHEME（主题色）时，调用 `resolve_theme_color` 固化为 RGB
- 返回 `Color(type="rgb", value="RRGGBB")`，丢弃主题色语义换保真
- 所有调用 `parse_color` 的地方需传入 presentation（parser 各模块）

### 步骤 5：修复文本解析（对齐 + 字号继承）
**文件**：[src/ppt_transfor/parser/text.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/text.py)

- `parse_paragraph`：alignment 为 None 且形状是 placeholder 时，调用 `resolve_placeholder_props` 获取继承的对齐方式
- `_parse_font`：font.size 为 None 时，从 `resolve_placeholder_props` 获取继承字号
- `_parse_font`：font.color 为 None 时，从 `resolve_placeholder_props` 获取继承颜色（已固化为 RGB）

### 步骤 6：修复 group 解析（chOff/chExt）
**文件**：[src/ppt_transfor/parser/group.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/group.py)

- 解析 group shape 的 `<p:grpSpPr>/<a:xfrm>` 下的 `chOff` 和 `chExt`
- 写入 Shape 模型的 `child_offset` 和 `child_extent` 字段
- 使用 `shape._element` 直接读 XML（python-pptx 未暴露 chOff/chExt API）

### 步骤 7：修复 group 渲染（回写 chOff/chExt）
**文件**：[src/ppt_transfor/renderer/group.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/renderer/group.py)

- 渲染 group 后，若模型有 `child_offset`/`child_extent`，回写到 `<p:grpSpPr>/<a:xfrm>` 的 `chOff`/`chExt`
- 使用 `group._element` 直接操作 XML

### 步骤 8：修复对比器（移除过度等价规则）
**文件**：[src/ppt_transfor/comparator/differ.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/comparator/differ.py)

移除/修改以下等价规则，让真实差异暴露：
1. **移除 `layout_name` 从 IGNORED_KEYS**（改为告警，因渲染统一用 Blank 仍会差异，但应可见）
2. **移除 alignment None≡CENTER 等价**（第 82-84 行）
3. **移除 None≡空容器等价**（第 89-90 行，保留空容器≡空容器）
4. **移除 placeholder≡text_box 等价**（第 102-103 行）
5. **移除 auto_shape≡text_box 等价**（第 104 行）
6. **保留 chart≡text_box 等价**（chart 确实不支持往返，这是已知限制）
7. **保留 _is_empty_text + text None≡空文本**（chart 降级场景仍需）

> 注：步骤 1-7 修复后，placeholder 仍会降级为 text_box（因渲染用 Blank 布局），但继承的对齐/字号/颜色已固化进 JSON，视觉差异会大幅减小。shape_type 差异保留可见，作为"降级"标记。

### 步骤 9：传递 presentation 到解析器各模块
**文件**：[src/ppt_transfor/parser/shape.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/shape.py)、[text.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/text.py)、[slide.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/slide.py)、[presentation.py](file:///i:/wefor/ppt-transfor/src/ppt_transfor/parser/presentation.py)

- `parse_presentation` 将 `prs` 对象传递到 `parse_slide` → `parse_shape` → `parse_text_frame` → `parse_color`
- 所有解析函数增加 `prs` 参数，用于主题色解析和 placeholder 继承解析

---

## 关键技术要点

### 1. 继承解析顺序
PPT 属性继承链：slide → layout → master → theme。解析时按此顺序向上查找，取第一个非 None 值。

### 2. placeholder 匹配
通过 `shape.placeholder_format.idx`（placeholder 索引）匹配 layout/master 中的 placeholder：
```python
idx = shape.placeholder_format.idx
layout_placeholder = slide.slide_layout.placeholders[idx]
```

### 3. 主题色固化
python-pptx 的 `color_format.theme_color` 返回 MSO_THEME_COLOR 枚举（如 ACCENT1）。需通过 theme part 的 clrScheme 解析为 RGB：
```python
# theme XML 中 <a:clrScheme> 下的 <a:accent1><a:srgbClr val="..."/>
```

### 4. group chOff/chExt
group shape 的 `<p:grpSpPr>/<a:xfrm>` 包含：
- `off/ext`：组合在幻灯片上的位置（已解析）
- `chOff/chExt`：子形状的坐标空间（本次新增）

### 5. SHAPE_TO_FIT_TEXT 溢出缓解
解析继承字号后，SHAPE_TO_FIT_TEXT 会基于正确字号计算尺寸，calendar 溢出问题应大幅缓解。剩余的字体替换导致的运行时差异，纯 JSON 无法完全解决（用户已确认接受）。

---

## 验证步骤

1. **单文件往返**：`uv run python main.py roundtrip input/cover_ending_page.pptx`
   - 检查 `out/json/cover_ending_page.json` 的 background 字段是否为黑色 RGB
   - 检查 title 段落的 alignment 是否为 CENTER
   - 检查 run 的 font.color 是否为浅灰色 RGB

2. **calendar 专项验证**：`uv run python main.py roundtrip input/content_page_component_calendar.pptx`
   - 检查 group 的 child_offset/child_extent 是否捕获
   - 检查日历单元格的 font.size 是否为继承的小字号
   - 用 PowerPoint 打开转换后 PPT，确认不再超出范围

3. **全量往返**：`uv run python main.py roundtrip-all`
   - 预期差异数会上升（因移除了过度等价规则，真实差异暴露）
   - 重点关注背景/对齐/颜色相关差异是否消除

4. **视觉抽检**：用 PowerPoint 打开几个转换后 PPT，与原始 PPT 对比
   - 背景色是否一致
   - title 是否居中
   - 辅色是否为浅灰

5. **测试**：`uv run pytest tests/`
   - 预期部分测试会失败（因对比器不再等价化）
   - 根据失败原因继续修复，直至视觉差异收敛

---

## 假设与约定

- **假设**：测试 PPT 的继承链均可通过 python-pptx 的 layout/master API 访问
- **假设**：主题色可通过 theme part 的 XML 解析为 RGB
- **约定**：主题色固化后 JSON 中颜色统一为 `type="rgb"`，不再保留 `type="theme"`
- **约定**：placeholder 降级为 text_box 是已知限制（渲染用 Blank 布局），但继承属性已固化进 JSON
- **约定**：calendar 溢出问题通过补 chOff/chExt + 继承字号缓解，剩余字体替换导致的运行时差异用户已接受
- **不引入**：图像对比层（用户已确认本次不引入）
