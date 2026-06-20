"""表格渲染器：Table 模型 → Table 形状。

add_table 后填充单元格，处理合并。
"""

from __future__ import annotations

from pptx.util import Emu

from ppt_transfor.models.schema import Shape
from ppt_transfor.renderer.text import render_text_frame


def render_table(slide, model: Shape):
    """渲染表格形状，返回 python-pptx GraphicFrame 对象。

    Args:
        slide: python-pptx Slide
        model: Shape 模型（shape_type == "table"）

    Returns:
        GraphicFrame 对象
    """
    tbl_model = model.table
    if tbl_model is None:
        return None

    rows = tbl_model.rows
    cols = tbl_model.cols

    graphic_frame = slide.shapes.add_table(
        rows,
        cols,
        Emu(model.left) if model.left is not None else None,
        Emu(model.top) if model.top is not None else None,
        Emu(model.width) if model.width is not None else None,
        Emu(model.height) if model.height is not None else None,
    )
    tbl = graphic_frame.table

    # 首行表头：显式设置，避免 add_table 默认 True
    try:
        tbl.first_row = tbl_model.first_row_header
    except Exception:
        pass

    # 填充单元格
    for row_idx in range(rows):
        for col_idx in range(cols):
            if row_idx >= len(tbl_model.cells) or col_idx >= len(tbl_model.cells[row_idx]):
                continue
            cell_model = tbl_model.cells[row_idx][col_idx]
            cell = tbl.cell(row_idx, col_idx)

            # 合并：span_x/span_y > 1 时需要 merge
            # 注意：python-pptx 的 merge 会将目标单元格合并到当前单元格
            if cell_model.span_x > 1 or cell_model.span_y > 1:
                try:
                    end_row = row_idx + cell_model.span_y - 1
                    end_col = col_idx + cell_model.span_x - 1
                    if end_row < rows and end_col < cols:
                        cell.merge(tbl.cell(end_row, end_col))
                except Exception:
                    pass

            # 文本
            render_text_frame(cell.text_frame, cell_model.text)

    return graphic_frame
