import logging
from typing import Tuple

import numpy as np
from openpyxl.cell import Cell
from openpyxl.cell.cell import TYPE_NUMERIC

from model.subject_model import Subjects

STU_DTYPE = np.dtype({'names': ['grade_name', 'class_name', 'name'] + [sub for sub, _ in Subjects.values()],
                      'formats': ['U32', 'U32', 'U32', 'f', 'f', 'f', 'f']})


class Student:

    def __init__(self, row: Tuple[Cell, ...]):
        # 年级 班级 学生姓名 语文 数学 英语 总评
        self.grade_name: str = row[0].value
        self.class_name: str = str(row[1].value)
        self.name: str = row[2].value
        self.chinese = float(row[3].value)
        self.math = float(row[4].value)
        self.two = self.chinese + self.math
        if len(row) == 5 or row[5] is None or row[5].value is None:
            self.english = float(0)
        else:
            self.english = float(row[5].value)

    @property
    def to_tuple(self):
        return self.grade_name, self.class_name, self.name, self.chinese, self.math, self.english, self.two


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
