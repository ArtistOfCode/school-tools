import logging
import os

from utils.utils import is_grade

XLSX = '.xlsx'


class Config:

    def __init__(self, root_dir):
        self.root_dir = root_dir
        # 及格分数
        self.pass_score = 60
        # 单科特优分数
        self.single_top_score = 92.5
        # 两科特优分数
        self.two_top_score = 185.0
        # 关爱人数比例
        self.care_rate = 0.2
        # 成绩数据目录
        self.data_path = f'{root_dir}/data/read'
        # 成绩分析结果
        self.result_path = f'{root_dir}/data/成绩分析结果.xlsx'
        # PPT模板
        self.ppt_template_path = f'{root_dir}/data/成绩分析模板.pptx'
        # PPT结果
        self.ppt_result_path = f'{root_dir}/data/成绩分析结果.pptx'
        # 是否生成关爱生列表
        self.need_care = False
        # 是否生成PPT
        self.need_ppt = False

    @property
    def file_data(self):
        file_paths = []
        data_paths = os.listdir(self.data_path)
        for file in data_paths:
            _file = os.path.join(self.data_path, file)
            if os.path.isfile(_file) and file.endswith(XLSX):
                _name = file.replace(XLSX, '')
                if is_grade(_name):
                    file_paths.append((_name, _file))
                else:
                    logging.error(f'文件名称必须是年级名称: {_name}')
                    continue
            else:
                continue
        file_paths.sort(key=lambda x: x[0])
        return file_paths
