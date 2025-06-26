import logging
import os
from typing import List

import numpy as np
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pptx import Presentation
from pptx.util import Inches

from model.config_model import Config
from model.score_model import ClassScore, SubjectScore
from model.student_model import is_valid_stu, STU_DTYPE, parse_stu
from utils.excel_utils import set_cell, set_title_cell, CellIndex, set_center_cell, set_float_cell
from utils.utils import Subjects, is_low_grade, is_school_class


class ScoreAnalyseService:

    def __init__(self, config: Config):
        self.config = config

    # 校级分析
    def school_analyse(self):
        # 1. 分析成绩结果
        school_result = []
        for title, file in self.config.file_data:
            logging.debug(f'加载成绩数据文件: {title, file}')
            workbook = load_workbook(file, True, False, True)
            school_result.append((title, self.grade_analyse(title, workbook)))
            workbook.close()
            logging.info(f'{title}分析完成')

        # 2. 保存分析结果
        workbook = Workbook()
        workbook.remove(workbook['Sheet'])
        for name, score in school_result:
            sheet = workbook.create_sheet(name)
            row = CellIndex()
            self.write_class(sheet, score, row)

            if not self.config.need_care:
                continue
            self.write_care_stu(sheet, score)
        workbook.save(self.config.result_path)
        logging.info('分析结果保存完成')

        # 3. 结果生成PPT
        pass

    # 校级分析
    def school_analyse2(self):
        school_workbook: Workbook = Workbook()
        school_workbook.remove(school_workbook['Sheet'])
        school_ppt = Presentation(self.config.ppt_template_path)
        for file_path in self.config.file_data:
            title = f'{os.path.basename(file_path)}'.replace('.xlsx', '')

            workbook: Workbook = load_workbook(file_path, True, False, True)
            school_score = self.grade_analyse(title, workbook)
            grade_sheet: Worksheet = school_workbook.create_sheet(title)

            grade_layout = school_ppt.slide_layouts[1]
            grade_slide = school_ppt.slides.add_slide(grade_layout)
            grade_slide.shapes.title.text = f'{title}成绩分析'

            row = CellIndex()
            self.write_class(grade_sheet, school_score, row)
            column = CellIndex(12)
            for subject in Subjects: self.write_care_stu(grade_sheet, school_score, subject, CellIndex(), column)
            for subject in Subjects: self.write_pptx(school_ppt, title, school_score, subject)

            workbook.close()
            logging.info(f'{title}分析完成！')
        school_workbook.save(self.config.result_path)
        school_ppt.save(self.config.ppt_result_path)
        logging.info('分析结果保存完成！')

    # 年级分析
    def grade_analyse(self, grade_name, workbook: Workbook):
        school_score = []
        # 第一遍循环分析：基本指标、一类关爱指标
        for class_score in self.class_analyse(grade_name, workbook):
            class_score.analyse1()
            school_score.append(class_score)
        # 第二遍循环分析二类关爱指标
        for class_score in school_score:
            class_score.analyse2(school_score[-1])
        return school_score

    # 班级分析
    def class_analyse(self, grade_name, workbook: Workbook):
        grade_students = []
        for sheetname in workbook.sheetnames:
            sheet = workbook[sheetname]
            class_students = [parse_stu(grade_name, r) for r in sheet.iter_rows(min_row=2) if
                              is_valid_stu(r, lambda: f'{grade_name} {sheetname}')]
            grade_students.extend(class_students)
            yield ClassScore(self.config, grade_name, sheetname, np.array(class_students, STU_DTYPE))
        yield ClassScore(self.config, grade_name, '校平', np.array(grade_students, STU_DTYPE))

    # 保存分析结果
    @staticmethod
    def write_class(sheet, score: List[ClassScore], row: CellIndex):
        for code, desc in [s.value for s in Subjects]:
            logging.debug(f'当前写入科目: {sheet.title} {desc}')
            _low = is_low_grade(sheet.title)
            if _low and code == Subjects.ENGLISH.code:
                continue

            # 定义表头
            care_score = getattr(score[0], code).care_stu_2[0]
            pass_name = '三科' if (not _low and code == Subjects.TWO.code) else ''
            care_name = f'率({care_score:g})' if _low else '平均分'
            headers = ['班级', '总人数', '平均分', f'{pass_name}及格人数', f'{pass_name}及格率', '特优人数', '特优率',
                       f'关爱{care_name}', '总评']

            # 写入科目名称
            set_title_cell(sheet.cell(row.value, 1), desc)
            row.next()

            # 写入成绩表头
            for _idx, header in enumerate(headers):
                set_title_cell(sheet.cell(row.value, _idx + 1), header)
            row.next()

            # 写入成绩
            for _score in score:
                logging.debug(f'当前写入班级: {sheet.title} {desc} {_score.name}')
                _sub_score: SubjectScore = getattr(_score, code)
                _idx = CellIndex(0)
                _cell = lambda: sheet.cell(row.value, _idx.next())

                set_cell(_cell(), _score.name)
                set_cell(_cell(), _score.total_stu)
                set_float_cell(_cell(), _sub_score.mean)
                set_cell(_cell(), _sub_score.pass_stu[0])
                set_float_cell(_cell(), _sub_score.pass_stu[1])
                if code != Subjects.ENGLISH.code:
                    set_cell(_cell(), _sub_score.top_stu[0])
                    set_float_cell(_cell(), _sub_score.top_stu[1])
                if is_low_grade:
                    set_float_cell(_cell(), _sub_score.care_stu_2[2])
                else:
                    set_float_cell(_cell(), _sub_score.care_stu_1[1])
                set_float_cell(_cell(), _sub_score.total)
                row.next()
            row.next()

    # 保存关爱学生
    @staticmethod
    def write_care_stu(sheet, score: List[ClassScore]):
        col = CellIndex(10)
        for code, desc in [s.value for s in Subjects]:
            logging.debug(f'当前写入关爱生科目: {sheet.title} {desc}')
            _low = is_low_grade(sheet.title)
            if _low and code == Subjects.ENGLISH.code:
                continue

            _row = CellIndex(0)
            col.next(3)

            for _score in score:
                logging.debug(f'当前写入关爱生班级: {sheet.title} {desc} {_score.name}')
                _sub_score: SubjectScore = getattr(_score, code)
                if is_school_class(_score.name) or _sub_score.care_stu_array.size == 0:
                    continue

                _col = CellIndex(col.value)

                # 写入表头
                set_title_cell(sheet.cell(_row.next(), _col.value),
                               f'{_score.name}班{desc}（{_sub_score.care_stu_array.size}）')
                set_title_cell(sheet.cell(_row.next(), _col.value), '姓名')

                _cell = lambda: sheet.cell(_row.value, _col.next())
                if code != Subjects.TWO.code:
                    set_title_cell(_cell(), '分数')
                else:
                    set_title_cell(_cell(), Subjects.CHINESE.desc)
                    set_title_cell(_cell(), Subjects.MATH.desc)
                    if _score.is_low:
                        set_title_cell(_cell(), Subjects.TWO.desc)
                    else:
                        set_title_cell(_cell(), Subjects.ENGLISH.desc)
                        set_title_cell(_cell(), Subjects.TWO.desc)
                _row.next()

                # 写入关爱生列表
                for stu in _sub_score.care_stu_array:
                    _col = CellIndex(col.value)

                    set_cell(sheet.cell(_row.value, _col.value), stu['name'])
                    if code != Subjects.TWO.code:
                        set_cell(_cell(), stu[code])
                    else:
                        set_cell(_cell(), stu[Subjects.CHINESE.code])
                        set_cell(_cell(), stu[Subjects.MATH.code])
                        if _score.is_low:
                            set_cell(_cell(), stu[Subjects.TWO.code])
                        else:
                            set_cell(_cell(), stu[Subjects.ENGLISH.code])
                            set_cell(_cell(), stu[Subjects.TWO.code])
                    _row.next()
                _row.next()

    # 保存分析结果PPT
    def write_pptx(self, school_ppt, title, school_score, subject: Subjects):
        low = title == '一年级' or title == '二年级'
        if low and subject == Subjects.ENGLISH: return
        subject_code, subject_name = subject.value

        # 添加幻灯片
        subject_layout = school_ppt.slide_layouts[2]
        class_layout = school_ppt.slide_layouts[3]
        subject_slide = school_ppt.slides.add_slide(subject_layout)
        subject_slide.shapes.title.text = f'{subject_name}情况分析'
        class_slide = school_ppt.slides.add_slide(class_layout)
        class_slide.shapes.title.text = f'{subject_name}情况分析'

        # 计算成绩表格表头
        if low:
            care_score = getattr(school_score[0], subject_code).care_stu_2[0]
            if subject == Subjects.CHINESE or subject == Subjects.MATH:
                headers = ['班级', '平均分', '及格率', f'关爱率\v{care_score}', '总评', '与校\v平差', '与区\v平差',
                           '名次', '教者']
            else:
                headers = ['班级', '平均分', '及格人数', '及格率', f'关爱率\v{care_score}', '总评', '与校\v平差',
                           '与区\v平差', '名次', '班主任']
        else:
            if subject == Subjects.CHINESE or subject == Subjects.MATH:
                headers = ['班级', '平均分', '及格率', '关爱\v平均分', '特优率', '总评', '与校\v平差', '与区\v平差',
                           '名次', '教者']
            elif subject == Subjects.ENGLISH:
                headers = ['班级', '平均分', '及格率', f'关爱\v平均分', '总评', '与校\v平差', '与区\v平差', '名次',
                           '教者']
            else:
                headers = ['班级', '平均分', '三科\v及格人数', f'三科\v及格率', '关爱\v平均分', '总评', '与校\v平差',
                           '与区\v平差', '名次', '班主任']

        # 成绩表格排版
        ppt_width = school_ppt.slide_width.inches
        # ppt_height = school_ppt.slide_height.inches
        max_row = len(school_score) + 2
        max_column = len(headers)
        width = Inches(1.2)
        height = Inches(0.5)
        left = Inches((ppt_width - len(headers) * 1.2) / 2)
        top = Inches(1.5)
        table = class_slide.shapes.add_table(max_row, max_column, left, top, width, height).table

        for idx, header in enumerate(headers):
            table.columns[idx].width = width
            set_center_cell(table.cell(0, idx), header)

        row = CellIndex()

        color = None
        for idx, class_score in enumerate(school_score):
            _subject: SubjectScore = getattr(class_score, subject_code)
            row_idx = row.value
            table.rows[row_idx].height = height

            if low:
                if subject == Subjects.CHINESE or subject == Subjects.MATH:
                    set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                    set_center_cell(table.cell(row_idx, 1), self.to_string(_subject.mean), color)
                    set_center_cell(table.cell(row_idx, 2), self.to_string(_subject.pass_stu[1]), color)
                    set_center_cell(table.cell(row_idx, 3), self.to_string(_subject.care_stu_2[2]), color)
                else:
                    set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                    set_center_cell(table.cell(row_idx, 1), self.to_string(_subject.mean), color)
                    set_center_cell(table.cell(row_idx, 2), self.to_string(_subject.pass_stu[0]), color)
                    set_center_cell(table.cell(row_idx, 3), self.to_string(_subject.pass_stu[1]), color)
                    set_center_cell(table.cell(row_idx, 4), self.to_string(_subject.care_stu_2[2]), color)
            else:
                if subject == Subjects.CHINESE or subject == Subjects.MATH:
                    set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                    set_center_cell(table.cell(row_idx, 1), self.to_string(_subject.mean), color)
                    set_center_cell(table.cell(row_idx, 2), self.to_string(_subject.pass_stu[1]), color)
                    set_center_cell(table.cell(row_idx, 3), self.to_string(_subject.care_stu_1[1]), color)
                    set_center_cell(table.cell(row_idx, 4), self.to_string(_subject.top_stu[1]), color)
                elif subject == Subjects.ENGLISH:
                    set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                    set_center_cell(table.cell(row_idx, 1), self.to_string(_subject.mean), color)
                    set_center_cell(table.cell(row_idx, 2), self.to_string(_subject.pass_stu[1]), color)
                    set_center_cell(table.cell(row_idx, 3), self.to_string(_subject.care_stu_1[1]), color)
                else:
                    set_center_cell(table.cell(row_idx, 0), class_score.name, color)
                    set_center_cell(table.cell(row_idx, 1), self.to_string(_subject.mean), color)
                    set_center_cell(table.cell(row_idx, 2), self.to_string(_subject.pass_stu[0]), color)
                    set_center_cell(table.cell(row_idx, 3), self.to_string(_subject.pass_stu[1]), color)
                    set_center_cell(table.cell(row_idx, 4), self.to_string(_subject.care_stu_1[1]), color)
            row.next()
        set_center_cell(table.cell(row.value, 0), '区平', color)
        row.next()

    @staticmethod
    def to_string(number):
        return str(round(number, 2))
