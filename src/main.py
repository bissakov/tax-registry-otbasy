import datetime
import logging
import warnings
from os import makedirs
from os.path import join
from typing import List, Optional, Tuple

from config import notifier, session
import colvir
from data_structures import DateRange, ReportInfo
from utils import kill_all_processes


def get_month_range(today: datetime.date) -> tuple:
    year = today.year if today.month != 1 else today.year - 1
    month = today.month - 1 if today.month != 1 else 12
    first_day = datetime.datetime(year, month, 1)

    if month == 12:
        last_day = datetime.datetime(year + 1, 1, 1) - datetime.timedelta(days=1)
    else:
        last_day = datetime.datetime(year, month + 1, 1) - datetime.timedelta(days=1)

    first_day_str = first_day.strftime('%d.%m.%y')
    last_day_str = last_day.strftime('%d.%m.%y')

    return first_day_str, last_day_str


def get_quarter_range(today: datetime.date) -> Optional[Tuple[str, str]]:
    if today.day != 15 and today.month not in [1, 4, 7, 10]:
        return
    if today.month == 1:
        first_day = datetime.date(today.year - 1, 10, 1)
        last_day = datetime.date(today.year - 1, 12, 31)
    else:
        first_day = datetime.date(today.year, today.month - 3, 1)
        last_day = datetime.date(today.year, today.month, 1) - datetime.timedelta(days=1)
    return (
        first_day.strftime('%d.%m.%y'),
        last_day.strftime('%d.%m.%y')
    )


def get_report_info(today: datetime.date) -> List[ReportInfo]:
    report_types = ['Z_160_DEPOFNO200_02', 'Z_160_DEPOFNO200_05', 'Z_160_DEPOFNO200']
    branches = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '11', '12',
                '13', '14', '15', '17', '18', '19', '20', '21', '26', '30', '31']
    russian_months = ['Январь', 'Февраль', 'Март', 'Апрель',
                      'Май', 'Июнь', 'Июль', 'Август',
                      'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    year = today.year if today.month != 1 else today.year - 1
    fserver_base_path = fr'\\fserver\ДБУИО_Новая\ДБУИО_Общая папка\ДБУ_Информация УНУ\{year}\ИПНи СН\200 по клиентам за {year}'
    local_base_path = r'C:\Users\robot.ad\Desktop\tax registry\reports'
    report_infos = []

    if today.day == 5:
        date_range = DateRange(*get_month_range(today))
        parent_report_path = russian_months[today.month - 2]
    else:
        date_range = DateRange(*get_quarter_range(today))
        parent_report_path = f'{(today.month - 1) // 3} квартал'

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
                range=date_range,
            ))
            makedirs(report_infos[-1].report_local_folder_path, exist_ok=True)
            makedirs(report_infos[-1].report_local_folder_path.replace('reports', 'converted_reports'),
                     exist_ok=True)
            # makedirs(report_infos[-1].report_fserver_folder_path, exist_ok=True)
    return report_infos


def main() -> None:
    logging.info('Start of the process')

    warnings.simplefilter(action='ignore', category=UserWarning)

    today = datetime.date.today()
    # today = datetime.date(2023, 2, 5)
    if not (today.day == 5 or (today.day == 15 and today.month in [1, 4, 7, 10])):
        return

    report_infos = get_report_info(today)

    kill_all_processes(proc_name='COLVIR')

    try:
        colvir.run_colvir(report_infos=report_infos)
    except RuntimeError as e:
        notifier.send_message(str(e))
        session.close()
        raise e
    except Exception as e:
        notifier.send_message(str(e))
        session.close()
        raise e

    logging.info('Successful end of the process')
    session.close()


main()
