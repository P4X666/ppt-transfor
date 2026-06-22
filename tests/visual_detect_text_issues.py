"""使用 Python 视觉批量扫描 PPT 文本可见性问题。

基于解析模型（已解析继承/主题色）检测低对比度文本，
并生成可视化诊断图，帮助定位黑底黑字/白底白字。

用法：
    uv run python tests/visual_detect_text_issues.py [input_dir] [output_dir]
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from ppt_transfor.comparator.vision_text import scan_all_pptx


def main() -> None:
    input_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("input")
    output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("out/vision_text")

    if not input_dir.exists():
        print(f"输入目录不存在: {input_dir}")
        sys.exit(1)

    print(f"扫描目录: {input_dir}")
    print(f"输出目录: {output_dir}")

    all_issues = scan_all_pptx(input_dir, output_dir)

    print("\n=== 扫描结果 ===")
    if not all_issues:
        print("✅ 未发现低对比度文本问题")
    else:
        total = 0
        for file_name, issues in all_issues.items():
            print(f"⚠️ {file_name}: {len(issues)} 处问题")
            total += len(issues)
            for issue in issues[:5]:
                print(
                    f"    slide={issue.slide_index + 1} shape={issue.shape_name} "
                    f"text='{issue.text[:30]}' text=#{issue.text_color or 'default'} bg=#{issue.bg_color} "
                    f"contrast={issue.contrast:.2f}"
                )
            if len(issues) > 5:
                print(f"    ... 还有 {len(issues) - 5} 处")
        print(f"\n共 {total} 处问题")

    print(f"\n可视化报告: {output_dir / 'report.png'}")


if __name__ == "__main__":
    main()
