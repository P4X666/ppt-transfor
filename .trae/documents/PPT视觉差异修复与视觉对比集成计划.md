# PPT 视觉差异修复与视觉对比集成计划

## 一、摘要

针对用户反馈的 4 类共性视觉差异（黑色背景丢失、title 居中丢失、浅灰色辅色丢失、calendar 溢出），通过代码探索已定位到 JSON 解析/渲染层的明确 bug。

**核心结论：这些问题不需要视觉识别作为"修复手段"，全部可在 JSON 层修复**；但视觉对比作为"验证手段"很有价值，因为 JSON 字段对比存在盲区（字号微差导致换行、形状重叠、视觉溢出 JSON 看不出）。

本计划分两部分：
1. **JSON 层 bug 修复**：主题色固化盲区、背景渲染容错、组合形状子坐标系精确复刻
2. **视觉对比工具集成**：LibreOffice + PyMuPDF + Pillow，作为客观验证手段

## 二、当前状态分析

### 2.1 已实现的保真机制
- `utils/inheritance.py`：`resolve_background`（背景继承链）、`resolve_placeholder_props`（placeholder 继承属性）、`resolve_theme_color`（主题色固化）
- `parser/slide.py`：调用 `resolve_background`
- `parser/shape.py`：对 placeholder 调用 `resolve_placeholder_props`
- `parser/text.py`：`_parse_font`/`parse_paragraph` 使用 `inherited_props` 兜底
- `utils/color.py`：`parse_color` 对 SCHEME 类型调用 `resolve_theme_color`
- `renderer/group.py`：`_apply_child_coords` 在添加子形状之后回写 chOff/chExt

### 2.2 已定位的 bug

#### Bug 1：主题色固化盲区（根因）
`inheritance.py` 内部调用 `parse_color` 时未传 `prs`，导致主题色无法固化为 RGB：

- `_get_fill_from_background`：`parse_color(fill.fore_color)` —— 背景色若为 SCHEME（如 BACKGROUND_1→黑）无法固化
- `_extract_font_props`：`parse_color(font.color)` —— placeholder 字体色若为 SCHEME 无法固化
- `_extract_defRPr_from_xml`：遇到 `<a:schemeClr val="..."/>` 直接保留为 `Color(type="theme", value=val)`，未调用主题色固化

**影响**：黑色背景（BACKGROUND_1 主题色）、浅灰色辅色（TEXT_2 主题色）在 JSON 中保留为 `Color(type="theme")`，渲染器 `apply_color` 对 theme 类型处理不可靠，导致颜色丢失。

#### Bug 2：背景渲染静默吞异常
`renderer/slide.py` 的 `_apply_background` 用 `try-except` 包裹所有逻辑，失败时静默不渲染，难以发现问题。

#### Bug 3：组合形状子坐标系未验证
`out/` 目录不存在，说明转换链路从未真正跑过产物。`renderer/group.py` 的 `_apply_child_coords` 调用顺序已调整（在添加子形状之后），但未实际验证。calendar.pptx 的溢出可能源于：
- chOff/chExt 回写方式问题
- 子形状坐标在子坐标系与 group 坐标系之间的变换缺失
- 字号继承错误导致 SHAPE_TO_FIT_TEXT 撑大形状

### 2.3 无视觉对比能力
- `pyproject.toml` 依赖仅 click/pydantic/python-pptx/rich，无任何视觉库
- `comparator/differ.py` 是纯 JSON 字段级 diff
- `out/` 目录不存在，从未生成过转换产物

## 三、修复方案

### 修复 1：主题色固化盲区（核心）— 部分已完成

**目标**：让背景色、placeholder 字体色的主题色（SCHEME）全部固化为 RGB。

**文件**：`src/ppt_transfor/utils/inheritance.py`

**已完成（执行阶段已修改）**：
- `_get_fill_from_background(bg_obj, prs=None)` 增加 `prs` 参数，调用 `parse_color(fill.fore_color, prs)`
- `resolve_background(slide, prs=None)` 增加 `prs` 参数，透传给 `_get_fill_from_background`
- `_extract_font_props(font, prs=None)` 增加 `prs` 参数，调用 `parse_color(font.color, prs)`
- `_extract_defRPr_from_xml(text_frame_element, prs=None)` 增加 `prs` 参数，遇到 schemeClr 时调用新增的 `_resolve_schemeclr_to_rgb(val, prs)` 固化为 RGB，失败降级保留 theme 类型
- 新增 `_resolve_schemeclr_to_rgb(scheme_val, prs)` 函数：从 schemeClr 的 val（如 'dk1', 'lt2'）直接查 clrScheme 固化为 RGB
- `resolve_placeholder_props(shape, slide, prs=None)` 增加 `prs` 参数，透传给 `_extract_defRPr_from_xml` 和 `_extract_font_props`

**待完成**：
- `parser/slide.py`：`parse_slide` 调用 `resolve_background(slide, prs)` 传入 prs（当前第 30 行未传 prs）
- `parser/shape.py`：第 189-193 行 `resolve_placeholder_props(shape, slide)` 改为 `resolve_placeholder_props(shape, slide, prs)`

**文件**：`src/ppt_transfor/utils/color.py`

`apply_color` 已支持 `Color(type="rgb")`（第 89-95 行），主题色固化后 JSON 中为 RGB 类型，可直接回写，无需改动。

### 修复 2：背景渲染容错改进

**目标**：背景渲染失败时输出警告，不再静默吞异常。

**文件**：`src/ppt_transfor/renderer/slide.py`

改动：`_apply_background` 保留 try-except（防止个别 slide 渲染失败中断整体流程），但增加 `logging.warning` 输出异常信息，便于定位问题。

```python
import logging
logger = logging.getLogger(__name__)

def _apply_background(slide, bg: Background | None) -> None:
    if bg is None or bg.type == "none":
        return
    try:
        fill = slide.background.fill
        if bg.type == "solid":
            fill.solid()
            if bg.color is not None:
                apply_color(fill.fore_color, bg.color)
    except Exception as e:
        logger.warning("背景渲染失败: %s, bg=%s", e, bg)
```

### 修复 3：组合形状子坐标系精确复刻

**目标**：确保 group 的 chOff/chExt 和子形状坐标精确复刻，calendar 不再溢出。

**文件**：`src/ppt_transfor/renderer/group.py`

当前 `_apply_child_coords` 已在添加子形状之后调用（正确）。需验证：
1. 子形状坐标是否需要按 chOff/chExt 变换
2. chOff/chExt 回写的 XML 结构是否正确

PowerPoint group 坐标变换原理：
- group 有两个坐标系：外部坐标系（ext，group 在 slide 中的位置/尺寸）和内部子坐标系（chOff/chExt）
- 子形状的 left/top/width/height 是子坐标系坐标
- PowerPoint 显示时按公式变换：`group_coord = chOff + (child_coord * ext / chExt)`

python-pptx 的 `GroupShapes` 添加子形状时，子形状坐标是相对于 group 左上角的坐标。若 chOff/chExt 已设置，PowerPoint 显示时会应用变换。

**实施步骤**：
1. 先实际运行 calendar.pptx 往返，观察 JSON 中 child_offset/child_extent 和子形状坐标
2. 对比原始 PPT 和转换后 PPT 的 group XML，定位差异
3. 若子形状坐标需要变换，在 `parser/group.py` 解析时把子形状坐标从子坐标系变换到 group 坐标系，渲染时直接用 group 坐标系坐标，chOff/chExt 设为 (0,0)/(ext)
4. 若 chOff/chExt 回写有问题，修正 `_apply_child_coords` 的 XML 生成逻辑

**文件**：`src/ppt_transfor/parser/group.py`（可能需改动）

若需坐标变换，新增 `_transform_child_coords(child, ch_off, ch_ext, group_ext)` 函数，把子形状坐标从子坐标系变换到 group 坐标系。

### 修复 4：引入视觉对比工具

**目标**：用 LibreOffice + PyMuPDF + Pillow 实现像素级视觉对比，作为客观验证手段。

**新增依赖**（`pyproject.toml`）：
```toml
dependencies = [
    "click>=8.4.1",
    "pydantic>=2.13.4",
    "python-pptx>=1.0.2",
    "rich>=15.0.0",
    "PyMuPDF>=1.24.0",   # PDF 转 PNG
    "Pillow>=10.0.0",    # 像素 diff
]
```

**系统依赖**：LibreOffice（提供 `soffice` 命令，用于 pptx → pdf）。Windows 下需安装 LibreOffice 并确保 `soffice.exe` 在 PATH，或通过路径调用。

**新增文件**：`src/ppt_transfor/comparator/visual.py`

核心函数：
```python
def pptx_to_images(pptx_path: Path, output_dir: Path) -> list[Path]:
    """PPTX → PDF（soffice）→ PNG 每页（PyMuPDF）"""
    # 1. soffice --headless --convert-to pdf --outdir <tmp> <pptx>
    # 2. fitz.open(pdf) 遍历 page，page.get_pixmap() 保存 PNG
    # 3. 返回 PNG 路径列表

def compare_images(img1: Path, img2: Path, tolerance: int = 10) -> dict:
    """像素级对比两张图片，返回差异度"""
    # Pillow ImageChops.difference，计算差异像素占比
    # 返回 {"diff_ratio": float, "max_diff": int}

def generate_diff_heatmap(img1: Path, img2: Path, output: Path) -> None:
    """生成差异热图"""
    # 差异区域红色高亮叠加

def visual_compare(original_pptx: Path, converted_pptx: Path, output_dir: Path) -> dict:
    """对比两个 PPTX 的视觉差异，逐页对比"""
    # 1. 两个 pptx 分别转 PNG
    # 2. 逐页 compare_images
    # 3. 生成差异热图
    # 4. 返回 {"pages": [...], "avg_diff_ratio": float}
```

**文件**：`main.py`

新增命令：
- `visual-compare <original_pptx> <converted_pptx> [--out]`：视觉对比两个 PPTX，输出差异报告 + 热图
- `roundtrip` 命令增加 `--visual` 选项：往返后自动做视觉对比
- `roundtrip-all` 命令增加 `--visual` 选项

**输出路径**：
- 视觉对比报告：`out/visual/<stem>_report.txt`
- 差异热图：`out/visual/<stem>_diff_page<N>.png`

## 四、假设与决策

### 4.1 假设
- LibreOffice 已安装或可安装（Windows 下需用户确认安装路径）
- 主题色固化后，`apply_color` 能正确处理 `Color(type="rgb")`（已验证 `utils/color.py` 第 89-95 行支持）
- placeholder 降级为 text_box 后，对齐属性已固化进 JSON，渲染器 `_apply_alignment` 能正确设置（需实际运行验证）

### 4.2 决策
1. **修复手段**：全部在 JSON 解析/渲染层，不引入 AI 视觉模型做修复
2. **验证手段**：引入 LibreOffice + PyMuPDF + Pillow 做像素级视觉对比
3. **calendar 溢出**：精确复刻子坐标系，不降级裁剪/缩放
4. **背景渲染**：保留 try-except 防中断，但加日志警告
5. **对比器等价规则**：暂不调整，先看修复后的实际差异

## 五、实施步骤

### 步骤 1：修复主题色固化盲区（部分已完成）
**已完成**：`inheritance.py` 的 5 个函数增加 prs 参数 + schemeClr 固化 + 新增 `_resolve_schemeclr_to_rgb`

**待完成**：
- 改 `parser/slide.py`：`parse_slide` 第 30 行 `resolve_background(slide)` 改为 `resolve_background(slide, prs)`
- 改 `parser/shape.py`：第 189-193 行 `resolve_placeholder_props(shape, slide)` 改为 `resolve_placeholder_props(shape, slide, prs)`

### 步骤 2：背景渲染容错改进
- 改 `renderer/slide.py`：`_apply_background` 增加 logging 警告

### 步骤 3：实际运行单文件往返验证
- 运行 `uv run python main.py roundtrip input/content_page_component_calendar.pptx`
- 运行 `uv run python main.py roundtrip input/cover_ending_page.pptx`（验证背景+对齐+颜色）
- 观察 JSON diff 报告，确认背景色、对齐、颜色是否固化

### 步骤 4：组合形状子坐标系精确复刻
- 根据 step 3 的 calendar 往返结果，定位 group 差异
- 若需坐标变换，改 `parser/group.py` 新增 `_transform_child_coords`
- 若 chOff/chExt 回写有问题，改 `renderer/group.py` 的 `_apply_child_coords`
- 重新运行 calendar 往返验证

### 步骤 5：引入视觉对比工具
- 改 `pyproject.toml`：新增 PyMuPDF + Pillow 依赖
- `uv sync` 安装依赖
- 新建 `comparator/visual.py`：实现 `pptx_to_images`、`compare_images`、`generate_diff_heatmap`、`visual_compare`
- 改 `main.py`：新增 `visual-compare` 命令，`roundtrip`/`roundtrip-all` 增加 `--visual` 选项

### 步骤 6：全量往返 + 视觉对比验证
- `uv run python main.py roundtrip-all --visual`
- 检查 `out/visual/` 下的差异报告和热图
- 对差异较大的页面，回到 JSON 层定位根因并修复

### 步骤 7：测试套件验证
- `uv run pytest tests/`
- 确保现有测试不回归

## 六、验证步骤

### 6.1 JSON 层验证
```bash
# 单文件往返（calendar）
uv run python main.py roundtrip input/content_page_component_calendar.pptx

# 全量往返
uv run python main.py roundtrip-all
```
预期：17 个文件 JSON diff 全部 PASS（或仅剩可接受的微小数值差异）。

### 6.2 视觉验证
```bash
# 单文件视觉对比
uv run python main.py visual-compare input/cover_ending_page.pptx out/pptx/cover_ending_page.pptx

# 全量视觉对比（通过 roundtrip-all --visual）
uv run python main.py roundtrip-all --visual
```
预期：差异热图显示背景、对齐、颜色基本一致；calendar 不再溢出。

### 6.3 测试套件
```bash
uv run pytest tests/ -v
```
预期：所有测试通过，无回归。

## 七、风险与回退

### 7.1 风险
- LibreOffice 未安装或 PATH 配置问题：视觉对比命令会失败，但不影响 JSON 层修复
- 主题色固化可能因 theme part 解析失败而降级：保留 theme 类型兜底，不会中断流程
- 组合形状坐标变换可能引入新差异：需逐文件验证

### 7.2 回退
- 所有改动均通过 git 管理，可随时回退
- 视觉对比工具是新增模块，不影响现有 parser/renderer 链路
- 主题色固化改动是参数透传，不改变原有调用逻辑

## 八、当前执行进度

- **已完成**：`inheritance.py` 主题色固化盲区修复（5 个函数增加 prs 参数 + 新增 `_resolve_schemeclr_to_rgb` + schemeClr 固化逻辑）
- **待完成**：步骤 1 剩余（parser/slide.py 和 parser/shape.py 调用处传 prs）、步骤 2-7 全部
