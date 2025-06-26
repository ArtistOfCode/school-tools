import logging
from os.path import dirname, abspath

from model.config_model import Config
from service.score_analyse import ScoreAnalyseService

ROOT_DIR = dirname(abspath(__file__))

logging.getLogger().setLevel(logging.DEBUG)
# noinspection SpellCheckingInspection
logging.basicConfig(format='%(asctime)s %(levelname)7s: %(message)s')

if __name__ == '__main__':
    logging.debug(f'项目根目录: {ROOT_DIR}')
    config = Config(ROOT_DIR)
    config.need_care = True
    ScoreAnalyseService(config).school_analyse()
