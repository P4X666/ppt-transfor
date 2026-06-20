"""往返测试：input 下每个 pptx 跑 roundtrip 并断言无差异。

首次运行预期会有差异，用于暴露 parser/renderer 缺陷，逐步修复。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from ppt_transfor.comparator.differ import diff_json
from ppt_transfor.parser.presentation import parse_presentation
from ppt_transfor.renderer.presentation import render_presentation

# 测试输入目录
INPUT_DIR = Path(__file__).parent.parent / "input"
# 测试输出目录
OUTPUT_DIR = Path(__file__).parent.parent / "out" / "test"


def _collect_pptx_files() -> list[Path]:
    """收集 input/ 下所有 .pptx 文件"""
    if not INPUT_DIR.exists():
        return []
    return sorted(INPUT_DIR.glob("*.pptx"))


@pytest.mark.parametrize(
    "pptx_path",
    _collect_pptx_files(),
    ids=lambda p: p.stem,
)
def test_roundtrip(pptx_path: Path, tmp_path: Path) -> None:
    """对单个 pptx 执行往返：解析 → 渲染 → 解析 → 对比

    断言差异列表为空（或在可接受范围内）。
    """
    # 1. 原始 PPT → 模型
    original_model = parse_presentation(pptx_path)

    # 2. 模型 → 转换后 PPT
    converted_pptx = tmp_path / f"{pptx_path.stem}_converted.pptx"
    render_presentation(original_model, converted_pptx)

    # 3. 转换后 PPT → 模型
    converted_model = parse_presentation(converted_pptx)

    # 4. 对比
    original_dict = original_model.model_dump(exclude_none=True)
    converted_dict = converted_model.model_dump(exclude_none=True)
    result = diff_json(original_dict, converted_dict)

    # 断言无差异（首次运行预期会失败，用于暴露缺陷）
    assert result.diff_count == 0, (
        f"{pptx_path.name} 往返存在 {result.diff_count} 处差异:\n"
        + "\n".join(f"  - {d}" for d in result.diffs[:20])
        + ("\n  ..." if len(result.diffs) > 20 else "")
    )
