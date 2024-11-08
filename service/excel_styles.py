from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

bold = Font(bold=True)
center = Alignment(horizontal='center')
thin_side = Side(style='thin', color='000000')
thin_border = Border(thin_side, thin_side, thin_side, thin_side)


def set_title_cell(cell: Cell, value):
    set_cell(cell, value, bold, center)


def set_cell(cell: Cell, value, font: Font = None, align: Alignment = None, fill: PatternFill = None,
             border: Border = thin_border):
    cell.value = value
    if font is not None: cell.font = font
    if align is not None: cell.alignment = align
    if fill is not None: cell.fill = fill
    if border is not None: cell.border = border
