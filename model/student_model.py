import logging

import numpy as np
from openpyxl.cell.cell import TYPE_NUMERIC

from utils.utils import is_low_grade, Subjects

STU_DTYPE = np.dtype({'names': ['grade_name', 'class_name', 'name'] + [s.code for s in Subjects],
                      'formats': ['U5'] * 3 + ['f'] * 4})

low_row = lambda row: len(row) == 5
high_row = lambda row: len(row) == 6
is_number = lambda cell: cell.data_type == TYPE_NUMERIC


def parse_stu(grade_name, row):
    _grade_name, *class_info = (r.value for r in row)

    if grade_name != _grade_name:
        logging.warning(f'文件名与年级不匹配: {grade_name} {_grade_name}')

    low_grade = is_low_grade(grade_name) and not low_row(row)
    if low_grade or (not is_low_grade(grade_name) and not high_row(row)):
        logging.warning(f'年级与成绩列不匹配: {grade_name} {_grade_name} {class_info}')

    english = 0
    if is_low_grade(grade_name) and len(row) == 5:
        class_name, name, chinese, math = class_info
    else:
        class_name, name, chinese, math, english = class_info

    logging.debug(
        f'加载学生成绩: {grade_name:<5s}{class_name:<6d}{name:<5s}\t{chinese:<.2f}\t{math:<.2f}\t{english:<.2f}')

    return grade_name, str(class_name), name, float(chinese), float(math), float(english), float(chinese) + float(math)


def is_valid_stu(row, name):
    _, _, _, chinese_cell, math_cell, *_ = row
    if not chinese_cell or not math_cell or not chinese_cell.value or not math_cell.value:
        logging.warning(f'该学生成绩忽略: {name()} {[r.value for r in row]}')
        return False
    if low_row(row) and (not is_number(chinese_cell) or not is_number(math_cell)):
        logging.warning(f'该学生成绩忽略: {name()} {[r.value for r in row]}')
        return False
    if high_row(row) and not is_number(row[5]):
        logging.warning(f'该学生成绩忽略: {name()} {[r.value for r in row]}')
        return False
    return True
