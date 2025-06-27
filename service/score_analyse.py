import logging

import numpy as np
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from model.config_model import Config
from model.score_model import ClassScore
from model.student_model import is_valid_stu, STU_DTYPE, parse_stu
from service.score_save import ScoreSave
from utils.utils import Subjects


class ScoreAnalyseService:

    def __init__(self, config: Config):
        self.config = config

    # 校级分析
    def school_analyse(self):
        # 1. 分析成绩结果
        school_result = []
        for grade, file in self.config.file_data:
            logging.debug(f'加载成绩数据文件: {grade, file}')
            excel = load_workbook(file, True, False, True)
            school_result.append((grade, self.grade_analyse(grade, excel)))
            excel.close()
            logging.info(f'{grade}分析完成')

        # 2. 保存分析结果
        save = ScoreSave(self.config)
        for grade, score in school_result:
            sheet = save.add_excel_sheet(grade, score)
            save.write_excel_score(sheet)
            save.write_excel_care(sheet)
            save.write_pptx()
        save.save()
        logging.info('分析结果保存完成')

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
        # 计算班级排名
        self.ranks_analyse(grade_name, school_score)
        return school_score

    # 班级分析
    def class_analyse(self, grade_name, workbook: Workbook):
        grade_stu = []
        for cls_name in workbook.sheetnames:
            sheet = workbook[cls_name]
            rows = sheet.iter_rows(min_row=2, values_only=True)
            cls_stu = [parse_stu(grade_name, r) for r in rows if is_valid_stu(r, lambda: f'{grade_name} {cls_name}')]
            grade_stu.extend(cls_stu)
            yield ClassScore(self.config, grade_name, cls_name, np.array(cls_stu, STU_DTYPE))
        yield ClassScore(self.config, grade_name, '校平', np.array(grade_stu, STU_DTYPE))

    # 计算班级排名
    @staticmethod
    def ranks_analyse(grade_name, school_score):
        for sub in Subjects:
            total = [getattr(s, sub.code).total for s in school_score[:-1]]
            ranks = np.argsort(-np.array(total)).argsort() + 1
            logging.debug(f'排名: {grade_name} {sub.desc} {ranks}')
            for i, s in enumerate(school_score[:-1]):
                getattr(s, sub.code).rank = int(ranks[i])
