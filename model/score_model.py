from typing import Optional

import numpy as np
from numpy import ndarray

from model.subject_model import Subjects

# 及格分数
PASS_SCORE = 60.0
# 单科特优分数
SINGLE_TOP_SCORE = 92.5
# 两科特优分数
TWO_TOP_SCORE = 185.0
# 关爱人数比例
CARE_RATE = 0.2


class SubjectScore:

    def __init__(self, subject: Subjects):
        # 科目 平均分 及格指标() 特优指标() 关爱指标(关爱人数，关爱率，关爱平均分)
        self.subject = subject
        # 平均分
        self.mean = 0.0
        # 及格指标（及格人数，及格率）
        self.pass_stu = (0, 0.0)
        # 特优指标（特优人数，特优率）
        self.top_stu = (0, 0.0)
        # 关爱指标1（关爱人数，关爱平均分）
        self.care_stu_1 = (0, 0.0)
        # 关爱指标2（关爱分数线，关爱人数，关爱率）
        self.care_stu_2 = (0.0, 0, 0.0)
        # 关爱学生列表
        self.care_stu_array: Optional[ndarray] = None

    def analyse(self, class_score: 'ClassScore'):
        chinese, _ = Subjects.CHINESE.value
        math, _ = Subjects.MATH.value
        english, _ = Subjects.ENGLISH.value
        two, _ = Subjects.TWO.value
        current, _ = self.subject.value

        total_stu = class_score.total_stu
        stu_array = class_score.array
        subject_array = stu_array[current]

        if self.subject != Subjects.TWO:
            # 单科成绩分析
            _mean = subject_array.mean()
            _pass_count = stu_array[subject_array >= PASS_SCORE].size
            _top_count = stu_array[subject_array >= SINGLE_TOP_SCORE].size
        else:
            # 总评成绩分析
            _chinese_pass = (stu_array[chinese] >= PASS_SCORE)
            _math_pass = (stu_array[math] >= PASS_SCORE)
            _english_pass = (stu_array[english] >= PASS_SCORE)

            _mean = subject_array.mean() / 2
            if class_score.is_low_grade:
                _pass_count = stu_array[_chinese_pass & _math_pass].size
            else:
                _pass_count = stu_array[_chinese_pass & _math_pass & _english_pass].size
            _top_count = stu_array[subject_array >= TWO_TOP_SCORE].size

        _pass_rate = _pass_count / total_stu * 100
        _top_rate = _top_count / total_stu * 100

        _care_count = int(total_stu * CARE_RATE)
        _care_array = np.sort(stu_array, order=current)[:_care_count][::-1]
        _subject_array = _care_array[current]
        _care_mean = _subject_array.mean()

        self.mean = self.round(_mean)
        self.pass_stu = _pass_count, self.round(_pass_rate)
        self.top_stu = _top_count, self.round(_top_rate)
        self.care_stu_1 = _care_count, self.round(_care_mean)
        self.care_stu_array = _care_array

        # 校平分析最后算出关爱分数线
        if class_score.is_school_class:
            self.care_stu_2 = _subject_array.max(), _care_count, self.round((total_stu - _care_count) / total_stu * 100)

    def analyse_care(self, class_score: 'ClassScore', school_score: 'SubjectScore'):
        if class_score.is_school_class: return
        current, _ = self.subject.value

        total_stu = class_score.total_stu
        stu_array = class_score.array
        subject_array = stu_array[current]

        _care_score, _, _ = school_score.care_stu_2
        _care_array = np.sort(stu_array[subject_array <= _care_score], order=current)[::-1]
        _care_count = _care_array.size

        self.care_stu_2 = _care_score, _care_count, self.round((total_stu - _care_count) / total_stu * 100)

        # 低年级使用二类关爱指标，重新赋值关爱学生列表
        if class_score.is_low_grade:
            self.care_stu_array = _care_array

    @staticmethod
    def round(num):
        return np.around(float(num), decimals=4)


class ClassScore:

    def __init__(self, grade_name, name: str, array: ndarray):
        # 年级名称 班级名称 总人数 语文成绩 数学成绩 英语成绩 总评成绩
        self.grade_name = grade_name
        self.name = name
        self.array = array
        self.total_stu = array.size
        self.chinese = SubjectScore(Subjects.CHINESE)
        self.math = SubjectScore(Subjects.MATH)
        self.english = SubjectScore(Subjects.ENGLISH)
        self.two = SubjectScore(Subjects.TWO)

    def analyse(self):
        for sub, _ in Subjects.values():
            subject: SubjectScore = getattr(self, sub)
            subject.analyse(self)

    def analyse_care(self, school_score: 'ClassScore'):
        for sub, _ in Subjects.values():
            subject: SubjectScore = getattr(self, sub)
            subject.analyse_care(self, getattr(school_score, sub))

    @property
    def is_low_grade(self):
        return self.grade_name == '一年级' or self.grade_name == '二年级'

    @property
    def is_school_class(self):
        return self.name == '校平'
