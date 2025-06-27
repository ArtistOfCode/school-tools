import logging
from pathlib import Path

from utils.utils import is_grade, GRADE


class Config:

    def __init__(self, root_dir):
        self.root_dir = Path(root_dir)
        # 及格分数
        self.pass_score = 60
        # 单科特优分数
        self.single_top_score = 92.5
        # 两科特优分数
        self.two_top_score = 185.0
        # 关爱人数比例
        self.care_rate = 0.2
        # 成绩数据目录
        self.data_path = self.root_dir / 'data/read'
        # 成绩分析结果
        self.result_path = self.root_dir / 'data/成绩分析结果.xlsx'
        # PPT模板
        self.ppt_template_path = self.root_dir / 'data/成绩分析模板.pptx'
        # PPT结果
        self.ppt_result_path = self.root_dir / 'data/成绩分析结果.pptx'
        # 是否生成关爱生列表
        self.need_care = False
        # 是否生成PPT
        self.need_ppt = False

    @property
    def file_data(self):
        file_paths = []
        for file in self.data_path.glob('*.xlsx'):
            if is_grade(n := file.stem):
                file_paths.append((n, file))
            else:
                logging.error(f'文件名称必须是年级名称: {n}')
        file_paths.sort(key=lambda x: GRADE.index(x[0]))
        return file_paths
