from decimal import Decimal
from enum import unique, Enum
from typing import Any


class SubjectInfo:

    def __init__(self, code: str, name: str, func: Any, factor: Decimal = Decimal('1')):
        # 科目编码 科目名称
        self.code = code
        self.name = name
        self.func = func
        self.factor = factor

    def __eq__(self, other):
        return self.code == other.code


@unique
class Subjects(Enum):
    CHINESE = SubjectInfo('chinese', '语文', lambda s: s.chinese)
    MATH = SubjectInfo('math', '数学', lambda s: s.math)
    ENGLISH = SubjectInfo('english', '英语', lambda s: s.english)
    TWO = SubjectInfo('two', '总评', lambda s: s.two, Decimal('2'))
