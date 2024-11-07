import logging
from decimal import Decimal
from typing import Tuple

from openpyxl.cell import Cell
from openpyxl.cell.cell import TYPE_NUMERIC


class Student:
    # 年级 班级 学生姓名 语文 数学 英语
    grade_name: str
    class_name: str
    name: str
    chinese: Decimal
    math: Decimal
    english: Decimal
    two: Decimal
    three: Decimal

    def __init__(self, row: Tuple[Cell, ...]):
        self.grade_name = row[0].value
        self.class_name = row[1].value
        self.name = row[2].value
        self.chinese = Decimal(str(row[3].value))
        self.math = Decimal(str(row[4].value))
        self.two = self.chinese + self.math
        if len(row) == 5 or row[5] is None or row[5].value is None:
            self.english = Decimal('0')
            self.three = Decimal('0')
        else:
            self.english = Decimal(str(row[5].value))
            self.three = self.two + self.english


def is_valid_stu(row: Tuple[Cell, ...]):
    chinese_cell, math_cell = row[3], row[4]
    if chinese_cell is None or math_cell is None:
        logging.warning(f'该学生成绩忽略:{[r.value for r in row]}')
        return False
    if chinese_cell.value is None or math_cell.value is None:
        logging.warning(f'该学生成绩忽略:{[r.value for r in row]}')
        return False
    if len(row) == 5 and (chinese_cell.data_type != TYPE_NUMERIC or math_cell.data_type != TYPE_NUMERIC):
        logging.warning(f'该学生成绩忽略:{[r.value for r in row]}')
        return False
    if len(row) == 6 and row[5].data_type != TYPE_NUMERIC:
        logging.warning(f'该学生成绩忽略:{[r.value for r in row]}')
        return False
    return True
