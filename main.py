import logging
from os.path import dirname

from service.score_analyse import ScoreAnalyseService

ROOT_DIR = dirname(__file__)

if __name__ == '__main__':
    logging.getLogger().setLevel(logging.INFO)
    # noinspection SpellCheckingInspection
    logging.basicConfig(format='%(asctime)s - %(levelname)7s: %(message)s')
    ScoreAnalyseService(ROOT_DIR).school_analyse()
