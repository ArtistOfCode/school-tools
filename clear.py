import logging
from os.path import dirname, abspath

from model.config_model import Config
from service.score_clear import ScoreClearService

ROOT_DIR = dirname(abspath(__file__))

logging.getLogger().setLevel(logging.DEBUG)
logging.basicConfig(format='%(asctime)s %(levelname)7s: %(message)s')

if __name__ == '__main__':
    logging.debug(f'项目根目录: {ROOT_DIR}')
    config = Config(ROOT_DIR)
    ScoreClearService(config).clear()
