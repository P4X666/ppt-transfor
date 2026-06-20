"""JSON 深度对比器。

对比两个 dict（原始 JSON vs 转换后 JSON），输出字段级差异列表。

忽略列表：往返必然变化的字段（shape_id、name、source_file、layout_name）。
- shape_id/name：重建后自动命名会变
- layout_name：渲染统一用 Blank 布局（设计决策），继承属性已固化进 JSON
浮点容差：浮点数差异 < 1e-6 视为相等。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

# 往返必然变化的字段，加入忽略列表避免噪声
IGNORED_KEYS: set[str] = {
    "shape_id",  # 重建后 ID 会变
    "name",  # 部分自动命名会变
    "source_file",  # 文件名可能不同
    "layout_name",  # 渲染统一用 Blank 布局（设计决策），继承属性已固化
}

# 浮点容差
FLOAT_TOLERANCE = 1e-6


@dataclass
class DiffItem:
    """单条差异"""

    path: str
    original: Any
    converted: Any

    def __str__(self) -> str:
        return f"{self.path}: {self.original!r} → {self.converted!r}"


@dataclass
class DiffResult:
    """对比结果"""

    diffs: list[DiffItem] = field(default_factory=list)

    @property
    def has_diff(self) -> bool:
        return len(self.diffs) > 0

    @property
    def diff_count(self) -> int:
        return len(self.diffs)

    def summary(self) -> str:
        """生成统计摘要"""
        if not self.diffs:
            return "无差异"
        # 按顶层路径分组统计
        groups: dict[str, int] = {}
        for d in self.diffs:
            top = d.path.split(".")[0].split("[")[0]
            groups[top] = groups.get(top, 0) + 1
        parts = [f"{k}: {v}" for k, v in sorted(groups.items())]
        return f"共 {self.diff_count} 处差异 | " + " | ".join(parts)


def _is_float(v: Any) -> bool:
    """判断是否为浮点数（非 int）"""
    return isinstance(v, float)


def _float_equal(a: float, b: float) -> bool:
    """浮点数相等判断（带容差）"""
    return abs(a - b) < FLOAT_TOLERANCE


def _values_equal(a: Any, b: Any) -> bool:
    """值相等判断（处理浮点容差与已知等价类型）"""
    # 都为 None
    if a is None and b is None:
        return True

    # alignment 等价：None（继承）与 CENTER（add_shape 默认主题对齐）等价
    # 这是 python-pptx 的固有行为差异：add_shape 创建的段落默认 CENTER 对齐
    # 注：placeholder 的对齐已固化进 JSON（不会是 None），此等价仅影响非 placeholder 形状
    if (a is None and b == "CENTER") or (b is None and a == "CENTER"):
        return True

    # 一方为 None：空容器视为等于 None（保留，避免噪声）
    if a is None or b is None:
        if (a is None and b in ([], {}, "", False)) or (b is None and a in ([], {}, "", False)):
            return True
        return False

    # 浮点
    if _is_float(a) or _is_float(b):
        try:
            return _float_equal(float(a), float(b))
        except (TypeError, ValueError):
            return a == b

    # shape_type 等价：
    # - chart 不支持往返，降级为 text_box 是已知限制（保留等价）
    # - placeholder 降级为 text_box：渲染用 Blank 布局无 placeholder，
    #   但继承的对齐/字号/颜色已固化进 JSON，视觉差异已消除（保留等价）
    # - auto_shape 降级为 text_box：无 auto_shape_type 的形状（如直线连接器）
    #   渲染时降级为 text_box，是已知限制（保留等价）
    if (a in ("chart", "auto_shape", "placeholder") and b == "text_box") or (
        b in ("chart", "auto_shape", "placeholder") and a == "text_box"
    ):
        return True

    # 容器类型交给上层递归处理
    return a == b


def _is_empty_text(text_val: Any) -> bool:
    """判断是否为空文本（chart 降级为 text_box 产生的默认空文本）。

    空文本定义：paragraphs 为空，或所有段落都没有非空 run。
    """
    if not isinstance(text_val, dict):
        return False
    paragraphs = text_val.get("paragraphs", [])
    if not paragraphs:
        return True
    for para in paragraphs:
        runs = para.get("runs", [])
        # 段落有非空 run 则非空文本
        for run in runs:
            if run.get("text"):
                return False
    return True


def _diff_recursive(
    original: Any,
    converted: Any,
    path: str,
    diffs: list[DiffItem],
) -> None:
    """递归对比两个值"""
    # text 字段：None 与空文本（chart 降级产生的默认空文本）等价
    if path.endswith(".text") or path == "text":
        if original is None and _is_empty_text(converted):
            return
        if converted is None and _is_empty_text(original):
            return

    # 类型不同（且都不是 None/空容器）
    if isinstance(original, dict) and isinstance(converted, dict):
        # dict 对比：遍历所有 key 的并集
        all_keys = set(original.keys()) | set(converted.keys())
        for key in sorted(all_keys):
            # 跳过忽略字段
            if key in IGNORED_KEYS:
                continue
            child_path = f"{path}.{key}" if path else key
            o_val = original.get(key)
            c_val = converted.get(key)
            _diff_recursive(o_val, c_val, child_path, diffs)
        return

    if isinstance(original, list) and isinstance(converted, list):
        # list 对比：按下标对比
        max_len = max(len(original), len(converted))
        for idx in range(max_len):
            child_path = f"{path}[{idx}]"
            o_val = original[idx] if idx < len(original) else None
            c_val = converted[idx] if idx < len(converted) else None
            _diff_recursive(o_val, c_val, child_path, diffs)
        return

    # 标量对比
    if not _values_equal(original, converted):
        diffs.append(DiffItem(path=path, original=original, converted=converted))


def diff_json(original: dict, converted: dict) -> DiffResult:
    """对比两个 JSON dict，返回 DiffResult。

    Args:
        original: 原始 JSON（dict）
        converted: 转换后 JSON（dict）

    Returns:
        DiffResult，包含差异列表与统计
    """
    result = DiffResult()
    _diff_recursive(original, converted, "", result.diffs)
    return result


def format_diff_report(result: DiffResult, title: str = "") -> str:
    """格式化差异报告为可读文本。

    Args:
        result: DiffResult
        title: 报告标题

    Returns:
        报告文本
    """
    lines: list[str] = []
    if title:
        lines.append(f"=== {title} ===")
    lines.append(result.summary())
    lines.append("")
    if result.diffs:
        lines.append("差异明细：")
        for d in result.diffs:
            lines.append(f"  - {d}")
    return "\n".join(lines)
