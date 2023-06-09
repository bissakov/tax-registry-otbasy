import datetime
import logging
from os import makedirs

from pywinauto import actionlogger


class LogFilter(logging.Filter):
    def filter(self, record):
        return 'Cannot retrieve text length for handle' not in record.getMessage()

makedirs(r'D:\Work\python_rpa\tax-registry-otbasy\logs', exist_ok=True)
actionlogger.enable()
logger = logging.getLogger()
logger.setLevel(logging.INFO)

formatter = logging.Formatter('%(asctime)s %(levelname)s %(name)s %(threadName)s : %(message)s')

file_handler = logging.FileHandler(
    f'../logs/{datetime.date.today().strftime("%d.%m.%y")}.log',
    encoding='utf-8'
)
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(formatter)

stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.INFO)
stream_handler.setFormatter(formatter)

logger.addFilter(LogFilter())
logger.addHandler(file_handler)
logger.addHandler(stream_handler)

for name in logging.root.manager.loggerDict:
    logger = logging.getLogger(name)
    logger.addFilter(LogFilter())
