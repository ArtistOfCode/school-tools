from decimal import Decimal

from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from model.score_model import SubjectInfo

bold = Font(bold=True)
center = Alignment(horizontal='center')
thin_side = Side(style='thin', color='000000')
thin_border = Border(thin_side, thin_side, thin_side, thin_side)


def set_title_cell(cell: Cell, value):
    set_cell(cell, value, bold, center)


def set_float_cell(cell: Cell, value):
    cell.number_format = '0.00'
    set_cell(cell, value)


def set_cell(cell: Cell, value, font: Font = None, align: Alignment = None, fill: PatternFill = None,
             border: Border = thin_border):
    cell.value = value
    if font is not None: cell.font = font
    if align is not None: cell.alignment = align
    if fill is not None: cell.fill = fill
    if border is not None: cell.border = border


class Row:
    value: int

    def __init__(self, row=1): self.value = row

    def next(self):
        self.value += 1


subjects = [
    SubjectInfo('chinese', '语文', lambda s: s.chinese),
    SubjectInfo('math', '数学', lambda s: s.math),
    SubjectInfo('english', '英语', lambda s: s.english),
    SubjectInfo('two', '两科', lambda s: s.two, Decimal('2')),
    SubjectInfo('three', '三科', lambda s: s.three, Decimal(3)),
]


def is_low_grade(name: str): return name == '一年级' or name == '二年级'
