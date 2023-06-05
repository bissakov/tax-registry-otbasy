import datetime
import os
import warnings
from os import makedirs
from os.path import join
from typing import List

import dotenv
import requests

from src import colvir
from src.bot_notification import TelegramNotifier
from src.data_structures import Credentials, DateRange, Process, ReportInfo
from src.logger import logger
from src.utils import kill_all_processes

logger = logger
session = requests.Session()
dotenv.load_dotenv()
notifier = TelegramNotifier(token=os.getenv('TOKEN_LOG'), chat_id=os.getenv(f'CHAT_ID_LOG'), session=session)
credentials = Credentials(usr=os.getenv(f'COLVIR_USR'), psw=os.getenv(f'COLVIR_PSW'))
process = Process(name='COLVIR', path=os.getenv('COLVIR_PROCESS_PATH'))


def get_month_range(today: datetime.date) -> tuple:
    year = today.year
    month = today.month - 1
    first_day = datetime.datetime(year, month, 1)

    if month == 12:
        last_day = datetime.datetime(year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        last_day = datetime.datetime(year, month + 1, 1) - datetime.timedelta(days=1)

    first_day_str = first_day.strftime('%d.%m.%y')
    last_day_str = last_day.strftime('%d.%m.%y')

    return first_day_str, last_day_str


def get_report_info() -> List[ReportInfo]:
    report_types = ['Z_160_DEPOFNO200_02', 'Z_160_DEPOFNO200_05', 'Z_160_DEPOFNO200']
    branches = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '11',
                '12', '13', '14', '15', '17', '18', '19', '20', '21', '26']
    russian_months = ['Январь', 'Февраль', 'Март', 'Апрель',
                      'Май', 'Июнь', 'Июль', 'Август',
                      'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    fserver_base_path = r'\\fserver\ДБУИО_Новая\ДБУИО_Общая папка\ДБУ_Информация УНУ\2023\ИПНи СН\200 по клиентам за 2023'
    local_base_path = r'C:\Users\robot.ad\Desktop\tax registry\reports'
    today = datetime.date.today()
    report_infos = []

    parent_report_path = russian_months[today.month - 2]

    for branch in branches:
        report_rel_folder_path = join(parent_report_path, branch)
        for report_type in report_types:
            report_name = f'{report_type}__{branch}.xls'
            report_rel_path = join(report_rel_folder_path, report_name)
            report_infos.append(ReportInfo(
                report_type=report_type,
                branch=branch,
                report_name=report_name,
                report_local_folder_path=join(local_base_path, report_rel_folder_path),
                report_local_full_path=join(local_base_path, report_rel_path),
                report_fserver_folder_path=join(fserver_base_path, report_rel_folder_path),
                report_fserver_full_path=join(fserver_base_path, report_rel_path),
                range=DateRange(*get_month_range(today)),
            ))
            makedirs(report_infos[-1].report_local_folder_path, exist_ok=True)
            makedirs(report_infos[-1].report_local_folder_path.replace('reports', 'converted_reports'),
                     exist_ok=True)
            # makedirs(report_infos[-1].report_fserver_folder_path, exist_ok=True)
    return report_infos


def main() -> None:
    logger.info('Start of the process')

    warnings.simplefilter(action='ignore', category=UserWarning)

    report_infos = get_report_info()

    kill_all_processes(proc_name='COLVIR')

    try:
        colvir.run_colvir(report_infos=report_infos)
    except RuntimeError as e:
        session.close()
        raise e
    except Exception as e:
        session.close()
        raise e

    logger.info('Successful end of the process')
    session.close()


if __name__ == '__main__':
    main()
