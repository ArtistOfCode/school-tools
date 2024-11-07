import decimal
from decimal import Decimal
from typing import List

from model.student_model import Student

PASS_SCORE = Decimal('60')
SINGLE_TOP_SCORE = Decimal('92.5')
TWO_TOP_SCORE = Decimal('185')
CARE_RATE = 0.2


class SubjectScore:
    # 总分数 及格人数 特优人数 关爱人数 及格率 特优率 关爱平均分 所有学生 关爱学生
    total_score: Decimal
    pass_count: int
    top_count: int
    care_count: int
    average_score: Decimal
    pass_rate: Decimal
    top_rate: Decimal
    care_score: Decimal
    total_stu: List[Student]
    care_stu: List[Student]

    def __init__(self):
        self.total_score = Decimal('0')
        self.pass_count = 0
        self.top_count = 0
        self.care_count = 0
        self.average_score = Decimal('0')
        self.pass_rate = Decimal('0')
        self.top_rate = Decimal('0')
        self.care_score = Decimal('0')
        self.total_stu = []
        self.care_stu = []


class ClassScore:
    # 班级名称 总人数 语文成绩 数学成绩 英语成绩 两科成绩 三科成绩
    name: str
    total_count: int
    chinese: SubjectScore
    math: SubjectScore
    english: SubjectScore
    two: SubjectScore
    three: SubjectScore

    def __init__(self, name: str):
        self.name = name
        self.total_count = 0
        self.chinese = SubjectScore()
        self.math = SubjectScore()
        self.english = SubjectScore()
        self.two = SubjectScore()
        self.three = SubjectScore()

    def add_class(self, class_score: 'ClassScore'):
        self.total_count += class_score.total_count
        self.count_class(self.chinese, class_score.chinese)
        self.count_class(self.math, class_score.math)
        self.count_class(self.english, class_score.english)
        self.count_class(self.two, class_score.two)
        self.count_class(self.three, class_score.three)

    @staticmethod
    def count_class(subject: SubjectScore, class_subject: SubjectScore):
        subject.total_score += class_subject.total_score
        subject.pass_count += class_subject.pass_count
        subject.top_count += class_subject.top_count
        subject.care_count += class_subject.care_count
        subject.total_stu.extend(class_subject.total_stu)

    def add_student(self, stu: Student):
        self.count_single_subject(self.chinese, stu.chinese)
        self.chinese.total_stu.append(stu)

        self.count_single_subject(self.math, stu.math)
        self.math.total_stu.append(stu)

        self.count_single_subject(self.english, stu.english)
        self.english.total_stu.append(stu)

        self.count_two_subject(self.two, stu)
        self.two.total_stu.append(stu)

        self.count_three_subject(self.three, stu)

    def count_single_subject(self, subject: SubjectScore, score: Decimal):
        subject.total_score += score
        if score >= PASS_SCORE: subject.pass_count += 1
        if score >= SINGLE_TOP_SCORE: subject.top_count += 1
        subject.care_count = int(self.total_count * CARE_RATE)

    def count_two_subject(self, subject: SubjectScore, stu: Student):
        subject.total_score += stu.two
        if stu.chinese >= PASS_SCORE and stu.math >= PASS_SCORE: subject.pass_count += 1
        if stu.two >= TWO_TOP_SCORE: subject.top_count += 1
        subject.care_count = int(self.total_count * CARE_RATE)

    def count_three_subject(self, subject: SubjectScore, stu: Student):
        subject.total_score += stu.three
        if stu.chinese >= PASS_SCORE and stu.math >= PASS_SCORE and stu.english >= PASS_SCORE: subject.pass_count += 1
        subject.care_count = int(self.total_count * CARE_RATE)

    def calc_subject(self, subject: SubjectScore, func, factor: Decimal = Decimal('1')):
        total_count = Decimal(self.total_count)
        subject.average_score = self.divide(subject.total_score / factor, total_count)
        subject.pass_rate = self.divide(Decimal(subject.pass_count) * 100, total_count)
        subject.top_rate = self.divide(Decimal(subject.top_count) * 100, total_count)

        subject.total_stu.sort(key=func)
        subject.care_stu = subject.total_stu[:subject.care_count]
        subject.care_stu.reverse()
        subject.care_score = self.average([func(s) for s in subject.care_stu])

    def calc_class(self):
        self.calc_subject(self.chinese, lambda s: s.chinese)
        self.calc_subject(self.math, lambda s: s.math)
        self.calc_subject(self.english, lambda s: s.english)
        self.calc_subject(self.two, lambda s: s.two, Decimal('2'))
        self.calc_subject(self.three, lambda s: s.three, Decimal('3'))

    @staticmethod
    def divide(n1: Decimal, n2: Decimal):
        return (n1 / n2).quantize(Decimal('0.0000'), decimal.ROUND_HALF_UP)

    def average(self, arr: List[Decimal]):
        if arr is None or len(arr) == 0: return Decimal('0')
        total_score = Decimal('0')
        for s in arr: total_score += s
        return self.divide(total_score, Decimal(len(arr)))
