import os
from typing import List

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from model.score_model import ClassScore
from model.student_model import is_valid_stu, Student


class ScoreAnalyseService:
    file_paths: List[str]
    result_path: str

    def __init__(self, root_dir):
        files = ['一年级.xlsx', '二年级.xlsx', '三年级.xlsx', '四年级.xlsx', '五年级.xlsx', '六年级.xlsx']
        self.file_paths = [f'{root_dir}/data/read/{file}' for file in files]
        self.result_path = f'{root_dir}/data/成绩分析结果.xlsx'

    # 校级分析
    def school_analyse(self):
        school_workbook: Workbook = Workbook()
        school_workbook.remove(school_workbook['Sheet'])
        for file_path in self.file_paths:
            workbook: Workbook = load_workbook(file_path, True, False, True)
            school_score = self.grade_analyse(workbook)
            grade_sheet: Worksheet = school_workbook.create_sheet(
                title=f'{os.path.basename(file_path)}'.replace('.xlsx', ''))
            self.write_class(grade_sheet, school_score, '语文', lambda s: s.chinese)
            self.write_class(grade_sheet, school_score, '数学', lambda s: s.math)
            self.write_class(grade_sheet, school_score, '英语', lambda s: s.english)
            self.write_class(grade_sheet, school_score, '两科', lambda s: s.two)
            self.write_class(grade_sheet, school_score, '三科', lambda s: s.three)
            workbook.close()
        school_workbook.save(self.result_path)

    # 年级分析
    def grade_analyse(self, workbook: Workbook):
        school_score: List[ClassScore] = []
        grade_score = ClassScore('校平')
        for sheetname in workbook.sheetnames:
            class_score = self.class_analyse(workbook[sheetname])
            school_score.append(class_score)
            grade_score.add_class(class_score)
        grade_score.calc_class()
        school_score.append(grade_score)
        return school_score

    # 班级分析
    @staticmethod
    def class_analyse(sheet: Worksheet):
        class_score = ClassScore(sheet.title)
        students = [Student(row) for row in list(filter(is_valid_stu, sheet.iter_rows(min_row=2)))]

        class_score.total_count = len(students)

        for student in students: class_score.add_student(student)
        class_score.calc_class()

        return class_score

    # 保存结果
    @staticmethod
    def write_class(grade_sheet: Worksheet, school_score: List[ClassScore], subject_name: str, func):
        grade_sheet.append([subject_name])
        grade_sheet.append(['班级', '总人数', '平均分', '及格人数', '及格率', '特优人数', '特优率', '关爱平均分'])
        for class_score in school_score:
            subject = func(class_score)
            grade_sheet.append(
                [class_score.name, class_score.total_count, subject.average_score, subject.pass_count,
                 subject.pass_rate, subject.top_count, subject.top_rate, subject.care_score])
        grade_sheet.append([])
