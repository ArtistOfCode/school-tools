from typing import Optional

import numpy as np
from numpy import ndarray

from model.config_model import Config
from utils.utils import is_low_grade, is_school_class, Subjects


class SubjectScore:

    def __init__(self, config: Config, subject: Subjects):
        # 科目
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
        # 总评
        self.total = 0.0
        # 与校平差
        self.diff = 0.0
        # 名次
        self.rank = 0
        # 全局配置
        self.config = config

    def analyse(self, class_score: 'ClassScore'):
        # 班级总人数
        _total = class_score.total_stu
        # 班级成绩数据
        _stu = class_score.array
        # 班级当前科成绩
        _sub = _stu[self.subject.code]

        _pass_score = self.config.pass_score

        if Subjects.is_single(self.subject):
            # 单科成绩分析
            _mean = _sub.mean()
            _pass = _stu[_sub >= _pass_score].size
            _top = _stu[_sub >= self.config.single_top_score].size
        else:
            # 校平成绩分析
            _chn_pass = (_stu[Subjects.CHN.code] >= _pass_score)
            _math_pass = (_stu[Subjects.MATH.code] >= _pass_score)
            _mean = _sub.mean() / 2
            if class_score.is_low:
                _pass = _stu[_chn_pass & _math_pass].size
            else:
                _eng_pass = (_stu[Subjects.ENG.code] >= _pass_score)
                _pass = _stu[_chn_pass & _math_pass & _eng_pass].size
            _top = _stu[_sub >= self.config.two_top_score].size

        _pass_rate = _pass / _total * 100
        _top_rate = _top / _total * 100

        # 计算出关爱人数
        _care = int(_total * self.config.care_rate)
        # 计算当前科关爱学生列表
        _care_arr = np.sort(_stu, order=self.subject.code)[:_care][::-1]
        # 计算当前科关爱平均分
        _sub_arr = _care_arr[self.subject.code]
        _care_mean = _sub_arr.mean()

        self.mean = self.round(_mean)
        self.pass_stu = _pass, self.round(_pass_rate)
        self.top_stu = _top, self.round(_top_rate)
        self.care_stu_1 = _care, self.round(_care_mean)
        self.care_stu_array = _care_arr
        # 高年级一类关爱指标计算总评
        if not class_score.is_low:
            if Subjects.is_english(self.subject):
                self.total = self.round(_mean * 0.4 + _pass_rate * 0.4 + _care_mean * 0.2)
            else:
                self.total = self.round(_mean * 0.4 + _pass_rate * 0.3 + _top_rate * 0.2 + _care_mean * 0.1)

        # 校平分析最后算出关爱分数线
        if class_score.is_school:
            self.care_stu_2 = _sub_arr.max(), _care, self.round((_total - _care) / _total * 100)
            if class_score.is_low and not Subjects.is_english(self.subject):
                self.total = self.round(_mean * 0.4 + _pass_rate * 0.4 + self.care_stu_2[2] * 0.2)

    def analyse_final(self, class_score: 'ClassScore', school_score: 'SubjectScore'):
        if class_score.is_school:
            return
        _total = class_score.total_stu
        _stu = class_score.array
        _sub = _stu[self.subject.code]

        _care_score, *_ = school_score.care_stu_2
        _care_arr = np.sort(_stu[_sub <= _care_score], order=self.subject.code)[::-1]
        _care = _care_arr.size

        # 计算二类关爱指标
        self.care_stu_2 = _care_score, _care, self.round((_total - _care) / _total * 100)

        # 低年级使用二类关爱指标，重新赋值关爱学生列表
        if class_score.is_low:
            self.care_stu_array = _care_arr
            # 计算总评
            self.total = self.round(self.mean * 0.4 + self.pass_stu[1] * 0.4 + self.care_stu_2[2] * 0.2)

        # 计算与校平差和名次
        self.diff = self.round(self.total - school_score.total)

    @staticmethod
    def round(num):
        return np.around(float(num), decimals=4)


class ClassScore:

    def __init__(self, config: Config, grade_name: str, name: str, array: ndarray):
        # 年级名称 班级名称 总人数 语文成绩 数学成绩 英语成绩 总评成绩
        self.grade_name = grade_name
        self.name = name
        self.array = array
        self.total_stu = array.size
        self.chinese = SubjectScore(config, Subjects.CHN)
        self.math = SubjectScore(config, Subjects.MATH)
        self.english = SubjectScore(config, Subjects.ENG)
        self.two = SubjectScore(config, Subjects.TWO)

    def analyse1(self):
        for sub in Subjects:
            subject: SubjectScore = getattr(self, sub.code)
            subject.analyse(self)

    def analyse2(self, school_score: 'ClassScore'):
        for sub in Subjects:
            subject: SubjectScore = getattr(self, sub.code)
            subject.analyse_final(self, getattr(school_score, sub.code))

    @property
    def is_low(self):
        return is_low_grade(self.grade_name)

    @property
    def is_school(self):
        return is_school_class(self.name)
