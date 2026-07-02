import logging

import numpy as np

from utils.utils import is_low_grade, Subjects

STU_DTYPE = np.dtype(
    {
        'names': ['grade_name', 'class_name', 'name'] + [s.code for s in Subjects],
        'formats': ['U5'] * 3 + ['f'] * 4,
    }
)


def parse_stu(grade_name, row):
    _grade_name, cls_name, name, chn, math, *eng = row

    if grade_name != _grade_name:
        logging.warning(f'文件名与年级不匹配: {grade_name} {_grade_name}')

    match len(row):
        case 5 if is_low_grade(grade_name):
            pass
        case 6 if not is_low_grade(grade_name):
            pass
        case _:
            logging.warning(f'年级与成绩列不匹配: {grade_name} {_grade_name} {row}')

    eng = eng[0] if eng else 0

    # logging.debug(f'加载学生成绩: {grade_name:<5s}{cls_name:<6d}{name:<5s}\t{chn:<.2f}\t{math:<.2f}\t{eng:<.2f}')
    return (
        grade_name,
        str(cls_name),
        name,
        float(chn),
        float(math),
        float(eng),
        float(chn) + float(math),
    )


def is_valid_stu(row, name):
    _, _, _, *score = row
    if not all([type(s) in (int, float) for s in score]):
        logging.warning(f'该学生成绩忽略: {name()} {row}')
        return False
    return True
