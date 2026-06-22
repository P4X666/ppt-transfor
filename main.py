"""PPT 转换服务 CLI 入口。

命令：
    parse <pptx_path> [--out out/json/xxx.json]      解析单个 PPT 为 JSON
    parse-all                                          解析 input/ 下所有 PPT
    render <json_path> [--out out/pptx/xxx.pptx]      从 JSON 渲染 PPT
    render-all                                         渲染 out/json/ 下所有 JSON
    compare <original_pptx> <converted_pptx>          对比原始与转换后 PPT（JSON 字段级）
    visual-compare <original_pptx> <converted_pptx>   视觉对比（像素级，需 LibreOffice）
    roundtrip <pptx_path> [--visual]                  完整往返：解析→渲染→对比
    roundtrip-all [--visual]                          对 input/ 下所有 PPT 执行往返
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

from ppt_transfor.comparator.differ import diff_json, format_diff_report
from ppt_transfor.models.schema import Presentation
from ppt_transfor.parser.presentation import parse_presentation
from ppt_transfor.renderer.presentation import render_presentation

console = Console()


def _ensure_dir(path: Path) -> None:
    """确保目录存在"""
    path.mkdir(parents=True, exist_ok=True)


def _model_to_json_file(model: Presentation, path: Path) -> None:
    """将 Presentation 模型写入 JSON 文件（UTF-8 无 BOM，缩进 2，保留中文）"""
    _ensure_dir(path.parent)
    json_str = model.model_dump_json(exclude_none=True, indent=2)
    path.write_text(json_str, encoding="utf-8")


def _model_from_json_file(path: Path) -> Presentation:
    """从 JSON 文件读取 Presentation 模型"""
    json_str = path.read_text(encoding="utf-8")
    return Presentation.model_validate_json(json_str)


def _default_json_path(pptx_path: Path) -> Path:
    """根据 pptx 路径推导默认 JSON 输出路径"""
    return Path("out/json") / f"{pptx_path.stem}.json"


def _default_pptx_path(json_path: Path) -> Path:
    """根据 json 路径推导默认 pptx 输出路径"""
    return Path("out/pptx") / f"{json_path.stem}.pptx"


def _default_report_path(name: str) -> Path:
    """默认对比报告路径"""
    return Path("out/compare") / f"{name}.txt"


def _default_visual_dir(name: str) -> Path:
    """默认视觉对比输出目录"""
    return Path("out/visual") / name


def _run_visual_compare(original_pptx: Path, converted_pptx: Path, name: str) -> None:
    """执行视觉对比并输出结果（供 roundtrip / roundtrip-all 复用）"""
    from ppt_transfor.comparator.visual import visual_compare

    visual_dir = _default_visual_dir(name)
    console.print(f"[cyan]视觉对比中:[/cyan] {original_pptx.name} vs {converted_pptx.name}")
    result = visual_compare(original_pptx, converted_pptx, visual_dir)

    if not result["success"]:
        console.print(f"[yellow]视觉对比跳过:[/yellow] {result['error']}")
        return

    # 输出逐页差异
    table = Table(title=f"视觉对比: {name}")
    table.add_column("页码", justify="right", style="cyan")
    table.add_column("差异比例", justify="right")
    table.add_column("最大差异", justify="right")
    table.add_column("尺寸一致", justify="center")

    for page in result["pages"]:
        ratio_pct = f"{page['diff_ratio'] * 100:.2f}%"
        size_ok = "是" if page.get("size_match", False) else "否"
        color = "green" if page["diff_ratio"] < 0.05 else ("yellow" if page["diff_ratio"] < 0.20 else "red")
        table.add_row(
            str(page["page"]),
            f"[{color}]{ratio_pct}[/{color}]",
            str(page["max_diff"]),
            size_ok,
        )

    console.print(table)
    console.print(
        f"\n平均差异: [{'green' if result['avg_diff_ratio'] < 0.05 else 'yellow'}]{result['avg_diff_ratio'] * 100:.2f}%[/{'green' if result['avg_diff_ratio'] < 0.05 else 'yellow'}] | "
        f"最大页差异: {result['max_page_diff'] * 100:.2f}%"
    )

    if result["heatmap_paths"]:
        console.print(f"差异热图已保存到: [cyan]{visual_dir}[/cyan]")

    # 保存报告
    report_path = visual_dir / "report.txt"
    lines = [f"视觉对比报告: {name}", f"平均差异比例: {result['avg_diff_ratio'] * 100:.2f}%", f"最大页差异: {result['max_page_diff'] * 100:.2f}%", ""]
    for page in result["pages"]:
        lines.append(f"页 {page['page']}: 差异 {page['diff_ratio'] * 100:.2f}%, 最大差异 {page['max_diff']}, 尺寸一致 {page.get('size_match', False)}")
    report_path.write_text("\n".join(lines), encoding="utf-8")


@click.group()
def cli() -> None:
    """PPT ↔ JSON 双向转换服务"""


@cli.command()
@click.argument("pptx_path", type=click.Path(exists=True, path_type=Path))
@click.option("--out", "out_path", type=click.Path(path_type=Path), default=None, help="输出 JSON 路径")
def parse(pptx_path: Path, out_path: Path | None) -> None:
    """解析单个 PPT 为 JSON"""
    if out_path is None:
        out_path = _default_json_path(pptx_path)

    console.print(f"[cyan]解析中:[/cyan] {pptx_path}")
    model = parse_presentation(pptx_path)
    _model_to_json_file(model, out_path)
    console.print(f"[green]已保存 JSON:[/green] {out_path}")
    console.print(f"  幻灯片数: {len(model.slides)}")
    total_shapes = sum(len(s.shapes) for s in model.slides)
    console.print(f"  形状总数: {total_shapes}")


@cli.command(name="parse-all")
def parse_all() -> None:
    """解析 input/ 下所有 PPT"""
    input_dir = Path("input")
    if not input_dir.exists():
        console.print("[red]input/ 目录不存在[/red]")
        sys.exit(1)

    pptx_files = sorted(input_dir.glob("*.pptx"))
    if not pptx_files:
        console.print("[yellow]input/ 下无 .pptx 文件[/yellow]")
        return

    table = Table(title="批量解析结果")
    table.add_column("文件", style="cyan")
    table.add_column("幻灯片数", justify="right")
    table.add_column("形状数", justify="right")
    table.add_column("状态", justify="center")

    for pptx_path in pptx_files:
        try:
            model = parse_presentation(pptx_path)
            out_path = _default_json_path(pptx_path)
            _model_to_json_file(model, out_path)
            total_shapes = sum(len(s.shapes) for s in model.slides)
            table.add_row(pptx_path.name, str(len(model.slides)), str(total_shapes), "[green]OK[/green]")
        except Exception as e:
            table.add_row(pptx_path.name, "-", "-", f"[red]FAIL: {e}[/red]")

    console.print(table)


@cli.command()
@click.argument("json_path", type=click.Path(exists=True, path_type=Path))
@click.option("--out", "out_path", type=click.Path(path_type=Path), default=None, help="输出 pptx 路径")
def render(json_path: Path, out_path: Path | None) -> None:
    """从 JSON 渲染 PPT"""
    if out_path is None:
        out_path = _default_pptx_path(json_path)

    console.print(f"[cyan]渲染中:[/cyan] {json_path}")
    model = _model_from_json_file(json_path)
    saved_path = render_presentation(model, out_path)
    console.print(f"[green]已保存 PPT:[/green] {saved_path}")


@cli.command(name="render-all")
def render_all() -> None:
    """渲染 out/json/ 下所有 JSON"""
    json_dir = Path("out/json")
    if not json_dir.exists():
        console.print("[red]out/json/ 目录不存在，请先执行 parse-all[/red]")
        sys.exit(1)

    json_files = sorted(json_dir.glob("*.json"))
    if not json_files:
        console.print("[yellow]out/json/ 下无 .json 文件[/yellow]")
        return

    for json_path in json_files:
        try:
            model = _model_from_json_file(json_path)
            out_path = _default_pptx_path(json_path)
            saved_path = render_presentation(model, out_path)
            console.print(f"[green]OK[/green] {json_path.name} → {saved_path}")
        except Exception as e:
            console.print(f"[red]FAIL[/red] {json_path.name}: {e}")


@cli.command()
@click.argument("original_pptx", type=click.Path(exists=True, path_type=Path))
@click.argument("converted_pptx", type=click.Path(exists=True, path_type=Path))
@click.option("--report", "report_path", type=click.Path(path_type=Path), default=None, help="对比报告输出路径")
def compare(original_pptx: Path, converted_pptx: Path, report_path: Path | None) -> None:
    """对比原始 PPT 与转换后 PPT（两者都解析为 JSON 后 diff）"""
    console.print(f"[cyan]解析原始 PPT:[/cyan] {original_pptx}")
    original_model = parse_presentation(original_pptx)
    original_dict = original_model.model_dump(exclude_none=True)

    console.print(f"[cyan]解析转换后 PPT:[/cyan] {converted_pptx}")
    converted_model = parse_presentation(converted_pptx)
    converted_dict = converted_model.model_dump(exclude_none=True)

    result = diff_json(original_dict, converted_dict)
    title = f"{original_pptx.name} vs {converted_pptx.name}"
    report = format_diff_report(result, title=title)
    console.print(report)

    if report_path is not None:
        _ensure_dir(report_path.parent)
        report_path.write_text(report, encoding="utf-8")
        console.print(f"\n[green]报告已保存:[/green] {report_path}")


@cli.command()
@click.argument("pptx_path", type=click.Path(exists=True, path_type=Path))
@click.option("--keep-intermediate", is_flag=True, default=True, help="保留中间 JSON 与 PPT 产物")
@click.option("--visual", is_flag=True, default=False, help="往返后执行视觉对比（需 LibreOffice）")
def roundtrip(pptx_path: Path, keep_intermediate: bool, visual: bool) -> None:
    """完整往返：解析 → 渲染 → 解析 → 对比"""
    json_path = _default_json_path(pptx_path)
    converted_pptx_path = _default_pptx_path(pptx_path)
    report_path = _default_report_path(pptx_path.stem)

    # 1. 原始 PPT → JSON
    console.print(f"[cyan]1. 解析原始 PPT:[/cyan] {pptx_path}")
    original_model = parse_presentation(pptx_path)
    _model_to_json_file(original_model, json_path)
    console.print(f"   [green]JSON 已保存:[/green] {json_path}")

    # 2. JSON → 转换后 PPT
    console.print(f"[cyan]2. 渲染转换后 PPT:[/cyan] {converted_pptx_path}")
    render_presentation(original_model, converted_pptx_path, source_pptx_path=pptx_path)
    console.print(f"   [green]PPT 已保存:[/green] {converted_pptx_path}")

    # 3. 转换后 PPT → JSON
    console.print(f"[cyan]3. 解析转换后 PPT[/cyan]")
    converted_model = parse_presentation(converted_pptx_path)

    # 4. 对比
    console.print(f"[cyan]4. 对比原始 JSON 与转换后 JSON[/cyan]")
    original_dict = original_model.model_dump(exclude_none=True)
    converted_dict = converted_model.model_dump(exclude_none=True)
    result = diff_json(original_dict, converted_dict)

    report = format_diff_report(result, title=f"往返对比: {pptx_path.name}")
    _ensure_dir(report_path.parent)
    report_path.write_text(report, encoding="utf-8")

    console.print(report)
    console.print(f"\n[green]报告已保存:[/green] {report_path}")

    # 5. 视觉对比（可选）
    if visual:
        console.print(f"[cyan]5. 视觉对比[/cyan]")
        _run_visual_compare(pptx_path, converted_pptx_path, pptx_path.stem)

    if not keep_intermediate:
        json_path.unlink(missing_ok=True)
        converted_pptx_path.unlink(missing_ok=True)


@cli.command(name="roundtrip-all")
@click.option("--visual", is_flag=True, default=False, help="往返后执行视觉对比（需 LibreOffice）")
def roundtrip_all(visual: bool) -> None:
    """对 input/ 下所有 PPT 执行往返测试"""
    input_dir = Path("input")
    if not input_dir.exists():
        console.print("[red]input/ 目录不存在[/red]")
        sys.exit(1)

    pptx_files = sorted(input_dir.glob("*.pptx"))
    if not pptx_files:
        console.print("[yellow]input/ 下无 .pptx 文件[/yellow]")
        return

    table = Table(title="批量往返对比结果")
    table.add_column("文件", style="cyan")
    table.add_column("差异总数", justify="right")
    table.add_column("视觉差异", justify="right")
    table.add_column("状态", justify="center")

    total_diffs = 0
    success_count = 0
    fail_count = 0

    for pptx_path in pptx_files:
        try:
            json_path = _default_json_path(pptx_path)
            converted_pptx_path = _default_pptx_path(pptx_path)
            report_path = _default_report_path(pptx_path.stem)

            original_model = parse_presentation(pptx_path)
            _model_to_json_file(original_model, json_path)
            render_presentation(original_model, converted_pptx_path, source_pptx_path=pptx_path)
            converted_model = parse_presentation(converted_pptx_path)

            original_dict = original_model.model_dump(exclude_none=True)
            converted_dict = converted_model.model_dump(exclude_none=True)
            result = diff_json(original_dict, converted_dict)

            report = format_diff_report(result, title=f"往返对比: {pptx_path.name}")
            _ensure_dir(report_path.parent)
            report_path.write_text(report, encoding="utf-8")

            total_diffs += result.diff_count
            if result.diff_count == 0:
                status = "[green]PASS[/green]"
                success_count += 1
            else:
                status = f"[yellow]DIFF[/yellow]"

            # 视觉对比
            visual_str = "-"
            if visual:
                from ppt_transfor.comparator.visual import visual_compare
                visual_dir = _default_visual_dir(pptx_path.stem)
                v_result = visual_compare(pptx_path, converted_pptx_path, visual_dir)
                if v_result["success"]:
                    visual_str = f"{v_result['avg_diff_ratio'] * 100:.2f}%"
                else:
                    visual_str = "跳过"

            table.add_row(pptx_path.name, str(result.diff_count), visual_str, status)
        except Exception as e:
            table.add_row(pptx_path.name, "-", "-", f"[red]FAIL: {e}[/red]")
            fail_count += 1

    console.print(table)
    console.print(f"\n汇总: 总差异 {total_diffs} 处 | PASS {success_count} | FAIL {fail_count}")
    console.print(f"详细报告见 [cyan]out/compare/[/cyan] 目录")
    if visual:
        console.print(f"视觉对比结果见 [cyan]out/visual/[/cyan] 目录")


@cli.command(name="visual-compare")
@click.argument("original_pptx", type=click.Path(exists=True, path_type=Path))
@click.argument("converted_pptx", type=click.Path(exists=True, path_type=Path))
@click.option("--out", "out_dir", type=click.Path(path_type=Path), default=None, help="输出目录")
def visual_compare_cmd(original_pptx: Path, converted_pptx: Path, out_dir: Path | None) -> None:
    """视觉对比两个 PPTX（像素级，需 LibreOffice）"""
    name = original_pptx.stem
    if out_dir is None:
        out_dir = _default_visual_dir(name)
    _run_visual_compare(original_pptx, converted_pptx, name)


if __name__ == "__main__":
    cli()
