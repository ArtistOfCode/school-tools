import logging
import os
from typing import List

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from model.score_model import ClassScore, SubjectInfo
from model.student_model import is_valid_stu, Student
from service.excel_styles import set_cell, set_title_cell, set_float_cell, CellIndex, is_low_grade, subjects


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

            row = CellIndex()
            for subject in subjects:
                if is_low_grade(title) and subject.is_high_subject(): continue
                self.write_class(grade_sheet, school_score, subject, row)

            column = CellIndex(14)
            for subject in subjects:
                if is_low_grade(title) and subject.is_high_subject(): continue
                self.write_care_stu(grade_sheet, school_score, subject, CellIndex(), column)

            workbook.close()
            logging.info(f'{title}分析完成！')
        school_workbook.save(self.result_path)
        logging.info('分析结果保存完成！')

    # 年级分析
    def grade_analyse(self, workbook: Workbook):
        school_score: List[ClassScore] = []
        grade_score = ClassScore('校平')
        for class_score in self.class_analyse(workbook):
            grade_score.add_class(class_score)
            school_score.append(class_score)
        grade_score.calc_class()
        school_score.append(grade_score)
        return school_score

    # 班级分析
    @staticmethod
    def class_analyse(workbook: Workbook):
        for sheetname in workbook.sheetnames:
            sheet = workbook[sheetname]
            class_score = ClassScore(sheetname)
            students = [Student(row) for row in list(filter(is_valid_stu, sheet.iter_rows(min_row=2)))]

            class_score.total_count = len(students)

            for student in students: class_score.add_student(student)
            class_score.calc_class()

            yield class_score

    # 保存结果
    @staticmethod
    def write_class(grade_sheet: Worksheet, school_score: List[ClassScore], subject_info: SubjectInfo, row: CellIndex):
        total_name = '三科' if not is_low_grade(grade_sheet.title) and subject_info.code == 'two' else ''
        headers = ['班级', '总人数', '平均分', f'{total_name}及格人数', f'{total_name}及格率', '特优人数', '特优率',
                   '关爱平均分', '总评', '与校平差', '排名']
        set_title_cell(grade_sheet.cell(row.value, 1), subject_info.name)
        row.next()

        for idx, header in enumerate(headers): set_title_cell(grade_sheet.cell(row.value, idx + 1), header)
        row.next()

        first_row = row.value
        c3 = None
        for class_score in school_score:
            subject = subject_info.func(class_score)
            row_index = row.value
            set_cell(grade_sheet.cell(row_index, 1), class_score.name)
            set_cell(grade_sheet.cell(row_index, 2), class_score.total_count)
            c1 = set_float_cell(grade_sheet.cell(row_index, 3), subject.average_score)
            set_cell(grade_sheet.cell(row_index, 4), subject.pass_count)
            c2 = set_float_cell(grade_sheet.cell(row_index, 5), subject.pass_rate)
            if subject_info.code != 'english':
                set_cell(grade_sheet.cell(row_index, 6), subject.top_count)
                c3 = set_float_cell(grade_sheet.cell(row_index, 7), subject.top_rate)
            c4 = set_float_cell(grade_sheet.cell(row_index, 8), subject.care_score)

            # 计算总评
            if subject_info.code == 'english':
                c5 = set_float_cell(grade_sheet.cell(row_index, 9),
                                    f'={c1.coordinate}*0.4+{c2.coordinate}*0.4+{c4.coordinate}*0.2')
            else:
                c5 = set_float_cell(grade_sheet.cell(row_index, 9),
                                    f'={c1.coordinate}*0.4+{c2.coordinate}*0.3+{c3.coordinate}*0.2+{c4.coordinate}*0.1')

            if c5 is not None and class_score.name != '校平':
                # 计算与校平差
                c6 = grade_sheet.cell(first_row + len(school_score) - 1, 9)
                set_float_cell(grade_sheet.cell(row_index, 10), f'={c5.coordinate}-{c6.coordinate}')
                # 计算排名
                c7 = grade_sheet.cell(first_row, 9)
                c8 = grade_sheet.cell(first_row + len(school_score) - 2, 9)
                set_cell(grade_sheet.cell(row_index, 11), f'=RANK({c5.coordinate},{c7.coordinate}:{c8.coordinate})')
            row.next()
        row.next()

    # 导出关爱学生
    @staticmethod
    def write_care_stu(grade_sheet: Worksheet, school_score: List[ClassScore], subject_info: SubjectInfo,
                       row: CellIndex, column: CellIndex):
        name_col, score_col = column.value, column.value + 1
        column.next(3)

        for class_score in school_score:
            subject = subject_info.func(class_score)
            if len(subject.care_stu) == 0 or subject.care_score == 0 or class_score.name == '校平': continue

            set_title_cell(grade_sheet.cell(row.value, name_col),
                           f'{class_score.name}班{subject_info.name}（{len(subject.care_stu)}）')
            row.next()

            set_title_cell(grade_sheet.cell(row.value, name_col), '姓名')
            set_title_cell(grade_sheet.cell(row.value, score_col), '分数')
            row.next()

            for stu in subject.care_stu:
                set_cell(grade_sheet.cell(row.value, name_col), stu.name)
                set_cell(grade_sheet.cell(row.value, score_col), subject_info.func(stu))
                row.next()
            row.next()
