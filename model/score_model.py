import decimal
from decimal import Decimal
from typing import List

from model.student_model import Student
from model.subject_model import SubjectInfo
from service.excel_styles import is_low_grade, subjects

# 及格分数
PASS_SCORE = Decimal('60')
# 单科特优分数
SINGLE_TOP_SCORE = Decimal('92.5')
# 两科特优分数
TWO_TOP_SCORE = Decimal('185')
# 关爱人数比例
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

    def calc_total(self, english=False):
        p4, p3, p2, p1 = Decimal('0.4'), Decimal('0.3'), Decimal('0.2'), Decimal('0.1')
        if english:
            return self.average_score * p4 + self.pass_rate * p4 + self.care_score * p2
        else:
            return self.average_score * p4 + self.pass_rate * p3 + self.top_rate * p2 + self.care_score * p1


class ClassScore:
    # 班级名称 总人数 语文成绩 数学成绩 英语成绩 总评成绩
    name: str
    total_count: int
    chinese: SubjectScore
    math: SubjectScore
    english: SubjectScore
    two: SubjectScore

    def __init__(self, name: str):
        self.name = name
        self.total_count = 0
        self.chinese = SubjectScore()
        self.math = SubjectScore()
        self.english = SubjectScore()
        self.two = SubjectScore()

    # 计算校平时添加班级
    def add_class(self, class_score: 'ClassScore'):
        self.total_count += class_score.total_count
        for subject in subjects: self.count_class(subject.func(self), subject.func(class_score))

    # 计算校平累计班级数据
    @staticmethod
    def count_class(subject: SubjectScore, class_subject: SubjectScore):
        subject.total_score += class_subject.total_score
        subject.pass_count += class_subject.pass_count
        subject.top_count += class_subject.top_count
        subject.care_count += class_subject.care_count
        subject.total_stu.extend(class_subject.total_stu)

    # 计算班级时添加学生
    def add_student(self, stu: Student):
        self.count_single_subject(self.chinese, stu.chinese)
        self.chinese.total_stu.append(stu)

        self.count_single_subject(self.math, stu.math)
        self.math.total_stu.append(stu)

        if not is_low_grade(stu.grade_name):
            self.count_single_subject(self.english, stu.english)
            self.english.total_stu.append(stu)

        self.count_two_subject(self.two, stu)
        self.two.total_stu.append(stu)

    # 计算班级时累计学生数据（单科）
    def count_single_subject(self, subject: SubjectScore, score: Decimal):
        subject.total_score += score
        if score >= PASS_SCORE: subject.pass_count += 1
        if score >= SINGLE_TOP_SCORE: subject.top_count += 1
        subject.care_count = int(self.total_count * CARE_RATE)

    # 计算班级时累计学生数据（总评）
    def count_two_subject(self, subject: SubjectScore, stu: Student):
        subject.total_score += stu.two
        if is_low_grade(stu.grade_name):
            if stu.chinese >= PASS_SCORE and stu.math >= PASS_SCORE: subject.pass_count += 1
        else:
            if stu.chinese >= PASS_SCORE and stu.math >= PASS_SCORE and stu.english >= PASS_SCORE: subject.pass_count += 1
        if stu.two >= TWO_TOP_SCORE: subject.top_count += 1
        subject.care_count = int(self.total_count * CARE_RATE)

    # 计算科目统计指标
    def calc_subject(self, subject_info: SubjectInfo):
        subject = subject_info.func(self)
        total_count = Decimal(self.total_count)

        subject.average_score = self.divide(subject.total_score / subject_info.factor, total_count)
        subject.pass_rate = self.divide(Decimal(subject.pass_count) * 100, total_count)
        subject.top_rate = self.divide(Decimal(subject.top_count) * 100, total_count)

        if len(subject.total_stu) > 0:
            subject.total_stu.sort(key=subject_info.func)
            subject.care_stu = subject.total_stu[:subject.care_count]
            subject.care_stu.reverse()
            subject.care_score = self.average([subject_info.func(s) for s in subject.care_stu])

    # 计算班级所有科目的统计指标（包括班级和校平）
    def calc_class(self):
        for subject in subjects: self.calc_subject(subject)

    @staticmethod
    def divide(n1: Decimal, n2: Decimal):
        return (n1 / n2).quantize(Decimal('0.0000'), decimal.ROUND_HALF_UP)

    def average(self, arr: List[Decimal]):
        if arr is None or len(arr) == 0: return Decimal('0')
        total_score = Decimal('0')
        for s in arr: total_score += s
        return self.divide(total_score, Decimal(len(arr)))
