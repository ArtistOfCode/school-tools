import decimal
from decimal import Decimal
from typing import List

from model.student_model import Student
from model.subject_model import SubjectInfo, Subjects
from service.excel_styles import is_low_grade

# 及格分数
PASS_SCORE = Decimal('60')
# 单科特优分数
SINGLE_TOP_SCORE = Decimal('92.5')
# 两科特优分数
TWO_TOP_SCORE = Decimal('185')
# 关爱人数比例
CARE_RATE = 0.2


class SubjectScore:

    def __init__(self, subject: Subjects):
        # 科目 总分数 及格人数 特优人数 关爱人数 及格率 特优率 关爱平均分 所有学生 关爱学生
        self.subject = subject
        self.total_score = Decimal('0')
        self.pass_count = 0
        self.top_count = 0
        self.care_count = 0
        self.average_score = Decimal('0')
        self.pass_rate = Decimal('0')
        self.top_rate = Decimal('0')
        self.care_score = Decimal('0')
        self.total_stu: List[Student] = []
        self.care_stu: List[Student] = []

    # 计算班级时累计学生数据
    def add_student(self, stu: Student):
        score: Decimal = getattr(stu, self.subject.value.code)
        self.total_score += score
        self.total_stu.append(stu)

        # 单科计算方法
        if self.subject != Subjects.TWO:
            if score >= PASS_SCORE: self.pass_count += 1
            if score >= SINGLE_TOP_SCORE: self.top_count += 1
        # 总评计算方法
        else:
            if is_low_grade(stu.grade_name) and stu.chinese >= PASS_SCORE and stu.math >= PASS_SCORE:
                self.pass_count += 1
            elif stu.chinese >= PASS_SCORE and stu.math >= PASS_SCORE and stu.english >= PASS_SCORE:
                self.pass_count += 1
            if score >= TWO_TOP_SCORE: self.top_count += 1

    # 计算校平累计班级数据
    def add_class(self, class_subject: 'SubjectScore'):
        self.total_score += class_subject.total_score
        self.pass_count += class_subject.pass_count
        self.top_count += class_subject.top_count
        self.care_count += class_subject.care_count
        self.total_stu.extend(class_subject.total_stu)

    def calc_total(self):
        p4, p3, p2, p1 = Decimal('0.4'), Decimal('0.3'), Decimal('0.2'), Decimal('0.1')
        if self.subject == Subjects.ENGLISH:
            return self.average_score * p4 + self.pass_rate * p4 + self.care_score * p2
        else:
            return self.average_score * p4 + self.pass_rate * p3 + self.top_rate * p2 + self.care_score * p1


class ClassScore:

    def __init__(self, name: str):
        # 班级名称 总人数 语文成绩 数学成绩 英语成绩 总评成绩
        self.name = name
        self.total_count = 0
        self.chinese = SubjectScore(Subjects.CHINESE)
        self.math = SubjectScore(Subjects.MATH)
        self.english = SubjectScore(Subjects.ENGLISH)
        self.two = SubjectScore(Subjects.TWO)

    # 计算校平时添加班级
    def add_class(self, class_score: 'ClassScore'):
        self.total_count += class_score.total_count
        for subject in Subjects:
            _subject: SubjectScore = getattr(self, subject.value.code)
            _subject.add_class(getattr(class_score, subject.value.code))

    # 计算班级时添加学生
    def add_student(self, stu: Student):
        care_count = int(self.total_count * CARE_RATE)

        self.chinese.add_student(stu)
        self.chinese.care_count = care_count

        self.math.add_student(stu)
        self.math.care_count = care_count

        if not is_low_grade(stu.grade_name):
            self.english.add_student(stu)
            self.english.care_count = care_count

        self.two.add_student(stu)
        self.two.care_count = care_count

    # 计算科目统计指标
    def calc_subject(self, subject: SubjectInfo):
        _subject: SubjectScore = getattr(self, subject.code)
        _total_count = Decimal(self.total_count)

        _subject.average_score = self.divide(_subject.total_score / subject.factor, _total_count)
        _subject.pass_rate = self.divide(Decimal(_subject.pass_count) * 100, _total_count)
        _subject.top_rate = self.divide(Decimal(_subject.top_count) * 100, _total_count)

        if len(_subject.total_stu) > 0:
            _subject.total_stu.sort(key=lambda s: getattr(s, subject.code))
            _subject.care_stu = _subject.total_stu[:_subject.care_count]
            _subject.care_stu.reverse()
            _subject.care_score = self.average([getattr(s, subject.code) for s in _subject.care_stu])

    # 计算班级所有科目的统计指标（包括班级和校平）
    def calc_class(self):
        for subject in Subjects: self.calc_subject(subject.value)

    @staticmethod
    def divide(n1: Decimal, n2: Decimal):
        return (n1 / n2).quantize(Decimal('0.0000'), decimal.ROUND_HALF_UP)

    @staticmethod
    def average(arr: List[Decimal]):
        if arr is None or len(arr) == 0: return Decimal('0')
        total_score = Decimal('0')
        for s in arr: total_score += s
        return ClassScore.divide(total_score, Decimal(len(arr)))
