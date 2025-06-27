from enum import unique, Enum

GRADE = ('一年级', '二年级', '三年级', '四年级', '五年级', '六年级')


def is_grade(grade_name):
    return grade_name in GRADE


def is_low_grade(grade_name):
    return grade_name in ('一年级', '二年级')


def is_school_class(name):
    return name == '校平'


@unique
class Subjects(Enum):
    CHN = 'chinese', '语文'
    MATH = 'math', '数学'
    ENG = 'english', '英语'
    TWO = 'two', '总评'

    def __init__(self, code, desc):
        self.code = code
        self.desc = desc

    @staticmethod
    def is_english(sub):
        return sub == Subjects.ENG or sub == Subjects.ENG.code

    @staticmethod
    def is_single(sub):
        return sub != Subjects.TWO or sub != Subjects.TWO.code

    @staticmethod
    def is_two(sub):
        return sub == Subjects.TWO or sub == Subjects.TWO.code
