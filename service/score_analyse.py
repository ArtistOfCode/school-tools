import os
from typing import List

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from model.score_model import ClassScore, SubjectInfo
from model.student_model import is_valid_stu, Student
from service.excel_styles import set_cell, set_title_cell, set_float_cell, Row, is_low_grade, subjects


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
            title = f'{os.path.basename(file_path)}'.replace('.xlsx', '')
            workbook: Workbook = load_workbook(file_path, True, False, True)
            school_score = self.grade_analyse(workbook)
            grade_sheet: Worksheet = school_workbook.create_sheet(title)
            row = Row()
            for subject in subjects:
                if is_low_grade(title) and subject.is_high_subject(): continue
                self.write_class(grade_sheet, school_score, subject, row)
            # for index, subject in enumerate(subjects):
            #     self.write_care_stu(grade_sheet, school_score, subject['name'], subject['func'], index)
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
    def write_class(grade_sheet: Worksheet, school_score: List[ClassScore], subject_info: SubjectInfo, row: Row):
        headers = ['班级', '总人数', '平均分', '及格人数', '及格率', '特优人数', '特优率', '关爱平均分']
        set_title_cell(grade_sheet.cell(row.value, 1), subject_info.name)
        row.next()

        for idx, header in enumerate(headers): set_title_cell(grade_sheet.cell(row.value, idx + 1), header)
        row.next()

        for class_score in school_score:
            subject = subject_info.func(class_score)
            row_index = row.value
            set_cell(grade_sheet.cell(row_index, 1), class_score.name)
            set_cell(grade_sheet.cell(row_index, 2), class_score.total_count)
            set_float_cell(grade_sheet.cell(row_index, 3), subject.average_score)
            set_cell(grade_sheet.cell(row_index, 4), subject.pass_count)
            set_float_cell(grade_sheet.cell(row_index, 5), subject.pass_rate)
            set_cell(grade_sheet.cell(row_index, 6), subject.top_count)
            set_float_cell(grade_sheet.cell(row_index, 7), subject.top_rate)
            set_float_cell(grade_sheet.cell(row_index, 8), subject.care_score)
            row.next()
        row.next()

    # 导出关爱学生
    @staticmethod
    def write_care_stu(grade_sheet: Worksheet, school_score: List[ClassScore], subject_name: str, func,
                       offset: int = 0):
        row_index, column_index = 1, (offset * 3) + 11
        name_col, score_col = column_index, column_index + 1

        for class_score in school_score:
            subject = func(class_score)
            if len(subject.care_stu) == 0 or subject.care_score == 0 or class_score.name == '校平': continue

            set_title_cell(grade_sheet.cell(row_index, name_col),
                           f'{class_score.name}班{subject_name}（{len(subject.care_stu)}）')
            row_index += 1

            set_title_cell(grade_sheet.cell(row_index, name_col), '姓名')
            set_title_cell(grade_sheet.cell(row_index, score_col), '分数')
            row_index += 1

            for stu in subject.care_stu:
                set_cell(grade_sheet.cell(row_index, name_col), stu.name)
                set_cell(grade_sheet.cell(row_index, score_col), func(stu))
                row_index += 1
            row_index += 1
