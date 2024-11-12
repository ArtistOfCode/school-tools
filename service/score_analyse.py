import decimal
import logging
import os
from decimal import Decimal
from typing import List

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pptx import Presentation
from pptx.util import Inches

from model.score_model import ClassScore, SubjectInfo, SubjectScore
from model.student_model import is_valid_stu, Student
from service.excel_styles import set_cell, set_title_cell, set_float_cell, CellIndex, is_low_grade, subjects, \
    set_center_cell


class ScoreAnalyseService:
    file_paths: List[str]
    result_path: str
    ppt_template_path: str
    ppt_result_path: str

    def __init__(self, root_dir):
        files = ['一年级.xlsx', '二年级.xlsx', '三年级.xlsx', '四年级.xlsx', '五年级.xlsx', '六年级.xlsx']
        self.file_paths = [f'{root_dir}/data/read/{file}' for file in files]
        self.result_path = f'{root_dir}/data/成绩分析结果.xlsx'
        self.ppt_template_path = f'{root_dir}/data/成绩分析模板.pptx'
        self.ppt_result_path = f'{root_dir}/data/成绩分析结果.pptx'

    # 校级分析
    def school_analyse(self):
        school_workbook: Workbook = Workbook()
        school_workbook.remove(school_workbook['Sheet'])
        school_ppt = Presentation(self.ppt_template_path)
        for file_path in self.file_paths:
            title = f'{os.path.basename(file_path)}'.replace('.xlsx', '')

            workbook: Workbook = load_workbook(file_path, True, False, True)
            school_score = self.grade_analyse(workbook)
            grade_sheet: Worksheet = school_workbook.create_sheet(title)

            grade_layout = school_ppt.slide_layouts[0]
            grade_slide = school_ppt.slides.add_slide(grade_layout)
            grade_slide.shapes.title.text = f'{title}成绩分析'

            row = CellIndex()
            for subject in subjects: self.write_class(grade_sheet, school_score, subject, row)
            column = CellIndex(15)
            for subject in subjects: self.write_care_stu(grade_sheet, school_score, subject, CellIndex(), column)
            for subject in subjects: self.write_pptx(school_ppt, title, school_score, subject)

            workbook.close()
            logging.info(f'{title}分析完成！')
        school_workbook.save(self.result_path)
        school_ppt.save(self.ppt_result_path)
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

    # 保存分析结果
    @staticmethod
    def write_class(grade_sheet: Worksheet, school_score: List[ClassScore], subject_info: SubjectInfo, row: CellIndex):
        if is_low_grade(grade_sheet.title) and subject_info.is_high_subject(): return

        pass_name = '三科' if not is_low_grade(grade_sheet.title) and subject_info.is_total_subject() else ''
        teacher_name = '班主任' if subject_info.is_total_subject() else '教者'

        headers = ['班级', '总人数', '平均分', f'{pass_name}及格人数', f'{pass_name}及格率', '特优人数', '特优率',
                   '关爱平均分', '总评', '与校平差', '排名', teacher_name]
        set_title_cell(grade_sheet.cell(row.value, 1), subject_info.name)
        row.next()

        for idx, header in enumerate(headers): set_title_cell(grade_sheet.cell(row.value, idx + 1), header)
        row.next()

        first_row = row.value
        c3 = None
        for idx, class_score in enumerate(school_score):
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
                # 添加教师
                set_cell(grade_sheet.cell(row_index, 12), f'{teacher_name}{idx + 1}')
            row.next()
        row.next()

    # 保存关爱学生
    @staticmethod
    def write_care_stu(grade_sheet: Worksheet, school_score: List[ClassScore], subject_info: SubjectInfo,
                       row: CellIndex, column: CellIndex):
        if is_low_grade(grade_sheet.title) and subject_info.is_high_subject(): return

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

    # 保存分析结果PPT
    def write_pptx(self, school_ppt, title, school_score, subject_info):
        if is_low_grade(title) and subject_info.is_high_subject(): return

        # 添加幻灯片
        subject_layout = school_ppt.slide_layouts[1]
        class_layout = school_ppt.slide_layouts[2]
        subject_slide = school_ppt.slides.add_slide(subject_layout)
        subject_slide.shapes.title.text = f'{subject_info.name}情况分析'
        class_slide = school_ppt.slides.add_slide(class_layout)
        class_slide.shapes.title.text = f'{subject_info.name}情况分析'

        # 计算成绩表格表头
        pass_name = '三科\v' if not is_low_grade(title) and subject_info.is_total_subject() else ''
        teacher_name = '班主任' if subject_info.is_total_subject() else '教者'
        headers = ['班级', '平均分', f'{pass_name}及格率', '特优率', '关爱\v平均分', '总评', '与校\v平差',
                   '名次', teacher_name]
        if subject_info.is_high_subject():
            headers = ['班级', '平均分', f'{pass_name}及格率', '关爱\v平均分', '总评', '与校\v平差', '名次',
                       teacher_name]
        if subject_info.is_total_subject():
            headers = ['班级', '平均分', f'{pass_name}及格人数', f'{pass_name}及格率', '特优率', '关爱\v平均分',
                       '总评', '与校\v平差', '名次', teacher_name]

        # 成绩表格排版
        ppt_width = school_ppt.slide_width.inches
        # ppt_height = school_ppt.slide_height.inches
        max_row = len(school_score) + 1
        max_column = len(headers)
        width = Inches(1.2)
        height = Inches(0.8)
        left = Inches((ppt_width - len(headers) * 1.2) / 2)
        top = Inches(1.5)
        table = class_slide.shapes.add_table(max_row, max_column, left, top, width, height).table

        for idx, header in enumerate(headers):
            table.columns[idx].width = width
            set_center_cell(table.cell(0, idx), header)

        row = CellIndex()

        # 计算总评
        for class_score in school_score:
            subject: SubjectScore = subject_info.func(class_score)
            if subject_info.is_high_subject():
                subject.total = subject.average_score * Decimal('0.4') + subject.pass_rate * Decimal(
                    '0.4') + subject.care_score * Decimal('0.2')
            else:
                subject.total = subject.average_score * Decimal('0.4') + subject.pass_rate * Decimal(
                    '0.3') + subject.top_rate * Decimal('0.2') + subject.care_score * Decimal('0.1')

        # 计算排名
        school_total = subject_info.func(school_score[len(school_score) - 1]).total
        sort_total = [subject_info.func(class_score).total for class_score in school_score][:len(school_score) - 1]
        sort_total.sort(reverse=True)

        for idx, class_score in enumerate(school_score):
            # color = 'ff0000' if class_score.name == '校平' else None
            color = None
            subject: SubjectScore = subject_info.func(class_score)
            row_idx = row.value
            if subject_info.is_high_subject():
                set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                set_center_cell(table.cell(row_idx, 1), self.to_string(subject.average_score), color)
                set_center_cell(table.cell(row_idx, 2), self.to_string(subject.pass_rate), color)
                set_center_cell(table.cell(row_idx, 3), self.to_string(subject.care_score), color)
                set_center_cell(table.cell(row_idx, 4), self.to_string(subject.total), color)
                if class_score.name != '校平':
                    set_center_cell(table.cell(row_idx, 5), self.to_string(subject.total - school_total), color)
                    set_center_cell(table.cell(row_idx, 6), str(sort_total.index(subject.total) + 1), color)
                    set_center_cell(table.cell(row_idx, 7), f'{teacher_name}{idx + 1}', color)
            elif subject_info.is_total_subject():
                set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                set_center_cell(table.cell(row_idx, 1), self.to_string(subject.average_score), color)
                set_center_cell(table.cell(row_idx, 2), str(subject.pass_count), color)
                set_center_cell(table.cell(row_idx, 3), self.to_string(subject.pass_rate), color)
                set_center_cell(table.cell(row_idx, 4), self.to_string(subject.top_rate), color)
                set_center_cell(table.cell(row_idx, 5), self.to_string(subject.care_score), color)
                set_center_cell(table.cell(row_idx, 6), self.to_string(subject.total), color)
                if class_score.name != '校平':
                    set_center_cell(table.cell(row_idx, 7), self.to_string(subject.total - school_total), color)
                    set_center_cell(table.cell(row_idx, 8), str(sort_total.index(subject.total) + 1), color)
                    set_center_cell(table.cell(row_idx, 9), f'{teacher_name}{idx + 1}', color)
            else:
                set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                set_center_cell(table.cell(row_idx, 1), self.to_string(subject.average_score), color)
                set_center_cell(table.cell(row_idx, 2), self.to_string(subject.pass_rate), color)
                set_center_cell(table.cell(row_idx, 3), self.to_string(subject.top_rate), color)
                set_center_cell(table.cell(row_idx, 4), self.to_string(subject.care_score), color)
                set_center_cell(table.cell(row_idx, 5), self.to_string(subject.total), color)
                if class_score.name != '校平':
                    set_center_cell(table.cell(row_idx, 6), self.to_string(subject.total - school_total), color)
                    set_center_cell(table.cell(row_idx, 7), str(sort_total.index(subject.total) + 1), color)
                    set_center_cell(table.cell(row_idx, 8), f'{teacher_name}{idx + 1}', color)
            row.next()

    @staticmethod
    def to_string(number: Decimal):
        return number.quantize(Decimal('0.00'), decimal.ROUND_HALF_UP).to_eng_string()
