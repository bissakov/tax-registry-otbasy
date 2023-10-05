import logging
import os
import sys
import traceback

from src import colvir, logger
from src.agent_initialization import system_paths
from src.bot_notification import send_message
from src.config import get_today
from src.report import get_report_info
from src.utils import kill_all_processes

logger.setup_logger()

if not str(os.getcwd()).endswith('Core_Agent'):
    sys.path = system_paths


def run() -> None:
    today = get_today()

    send_message(message=f'Start of the process for {today}')
    logging.info('Start of the process')
    if not (today.day == 5 or (today.day == 15 and today.month in [1, 4, 7, 10])):
        return

    report_infos = get_report_info(today=today)

    kill_all_processes(proc_name='COLVIR')
    kill_all_processes(proc_name='EXCEL')

    try:
        colvir.run_colvir(report_infos=report_infos)
    except Exception as error:
        error_msg = traceback.format_exc()
        send_message(f'Error occured on robot-21\n'
                     f'Process: Tax Registry'
                     f'{error_msg}')
        raise error
    logging.info('Successful end of the process')
