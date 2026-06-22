"""表格解析器：Table 形状 → Table 模型。

遍历单元格，记录合并（span_x/span_y）与文本。
"""

from __future__ import annotations

from ppt_transfor.models.schema import TableCell, Table
from ppt_transfor.parser.text import parse_text_frame
from ppt_transfor.utils.inheritance import extract_txbody_default_props


def parse_table(shape, prs=None) -> Table:
    """解析表格形状

    Args:
        shape: python-pptx GraphicFrame（含表格）
        prs: 所属 Presentation（用于主题色固化）
    """
    tbl = shape.table
    rows = len(tbl.rows)
    cols = len(tbl.columns)

    model = Table(rows=rows, cols=cols)

    # 首行表头标记（通过样式推断，简化处理）
    model.first_row_header = bool(getattr(tbl, "first_row", False))

    # 单元格矩阵
    cells = []
    for row_idx in range(rows):
        row_cells = []
        for col_idx in range(cols):
            cell = tbl.cell(row_idx, col_idx)
            cell_model = TableCell()

            # 文本：传入单元格默认样式，避免对齐/字号丢失
            cell_defaults = extract_txbody_default_props(cell.text_frame._element, prs)
            cell_model.text = parse_text_frame(cell.text_frame, prs, cell_defaults)

            # 合并信息：span_x 为横向合并数，span_y 为纵向合并数
            # python-pptx 通过 cell.span_x / cell.span_y 获取
            try:
                cell_model.span_x = int(cell.span_x)
            except Exception:
                cell_model.span_x = 1
            try:
                cell_model.span_y = int(cell.span_y)
            except Exception:
                cell_model.span_y = 1

            row_cells.append(cell_model)
        cells.append(row_cells)

    model.cells = cells
    return model
