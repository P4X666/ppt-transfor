"""对所有 input PPTX 执行往返转换并做视觉对比，输出差异报告。"""

from __future__ import annotations

import json
import sys
from pathlib import Path

# 把 src 加入路径
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from ppt_transfor.comparator.visual import visual_compare
from ppt_transfor.renderer.presentation import render_presentation
from ppt_transfor.parser.presentation import parse_presentation


def main() -> None:
    input_dir = Path("input")
    out_dir = Path("out")
    pptx_dir = out_dir / "pptx"
    visual_dir = out_dir / "visual_scan"
    pptx_dir.mkdir(parents=True, exist_ok=True)
    visual_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(input_dir.glob("*.pptx"))
    # 排除临时文件
    files = [f for f in files if not f.name.startswith("~$")]

    results = []
    for pptx in files:
        print(f"处理: {pptx.name}")
        converted = pptx_dir / pptx.name
        try:
            prs = parse_presentation(str(pptx))
            render_presentation(prs, str(converted))
        except Exception as e:
            print(f"  转换失败: {e}")
            results.append({"file": pptx.name, "error": str(e)})
            continue

        output_subdir = visual_dir / pptx.stem
        try:
            report = visual_compare(pptx, converted, output_subdir)
        except Exception as e:
            print(f"  视觉对比失败: {e}")
            results.append({"file": pptx.name, "error": str(e)})
            continue

        max_diff = report["max_page_diff"]
        avg_diff = report["avg_diff_ratio"]
        results.append({
            "file": pptx.name,
            "max_page_diff": max_diff,
            "avg_diff_ratio": avg_diff,
            "pages": report["pages"],
            "heatmaps": [str(p) for p in report["heatmap_paths"]],
            "success": report["success"],
            "error": report["error"],
        })
        print(f"  max_diff={max_diff:.4f}, avg_diff={avg_diff:.4f}")

    # 按差异从大到小排序
    results.sort(key=lambda x: x.get("max_page_diff", 0), reverse=True)

    report_path = visual_dir / "report.json"
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print("\n=== 视觉差异报告 ===")
    for r in results:
        if "error" in r:
            print(f"{r['file']}: 错误 {r['error']}")
        else:
            flag = "⚠️" if r["max_page_diff"] > 0.01 else "✅"
            print(f"{flag} {r['file']}: max={r['max_page_diff']:.4f}, avg={r['avg_diff_ratio']:.4f}")
    print(f"\n详细报告: {report_path}")


if __name__ == "__main__":
    main()
