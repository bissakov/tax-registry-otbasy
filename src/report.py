from datetime import date, datetime, timedelta
from os import makedirs
from os.path import dirname, join
from typing import Optional

from src.config import get_fserver_path, get_local_path
from src.data_structures import DateRange, ReportInfo


def get_month_range(today: datetime.date) -> tuple:
    year = today.year if today.month != 1 else today.year - 1
    month = today.month - 1 if today.month != 1 else 12
    first_day = datetime(year, month, 1)

    if month == 12:
        last_day = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        last_day = datetime(year, month + 1, 1) - timedelta(days=1)

    first_day_str = first_day.strftime('%d.%m.%y')
    last_day_str = last_day.strftime('%d.%m.%y')

    return first_day_str, last_day_str


def get_quarter_range(today: datetime.date) -> Optional[tuple[str, str]]:
    if today.day != 15 and today.month not in [1, 4, 7, 10]:
        return
    if today.month == 1:
        first_day = date(today.year - 1, 10, 1)
        last_day = date(today.year - 1, 12, 31)
    else:
        first_day = date(today.year, today.month - 3, 1)
        last_day = date(today.year, today.month, 1) - timedelta(days=1)
    return (
        first_day.strftime('%d.%m.%y'),
        last_day.strftime('%d.%m.%y')
    )


def get_report_info(today: datetime.date) -> list[ReportInfo]:
    reports_path = get_local_path()
    fserver_path = get_fserver_path()

    report_types = ['Z_160_DEPOFNO200_02', 'Z_160_DEPOFNO200_05', 'Z_160_DEPOFNO200']
    branches = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '11', '12',
                '13', '14', '15', '17', '18', '19', '20', '21', '26', '30', '31']
    russian_months = ['Январь', 'Февраль', 'Март', 'Апрель',
                      'Май', 'Июнь', 'Июль', 'Август',
                      'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']

    report_infos = []

    if today.day == 5:
        date_range = DateRange(*get_month_range(today=today))
        parent_report_path = russian_months[today.month - 2]
    else:
        date_range = DateRange(*get_quarter_range(today=today))
        parent_report_path = f'{(today.month - 1) // 3} квартал'

    for branch in branches:
        report_rel_folder_path = join(parent_report_path, branch)
        for report_type in report_types:
            report_name = f'{report_type}__{branch}.xls'
            report_rel_path = join(report_rel_folder_path, report_name)
            report_infos.append(ReportInfo(
                report_type=report_type,
                branch=branch,
                local_full_path=join(reports_path, report_rel_path),
                xlsb_full_path=join(reports_path, report_rel_path).replace('.xls', '.xlsb'),
                fserver_full_path=join(fserver_path, report_rel_path).replace('.xls', '.xlsb'),
                range=date_range,
            ))
            makedirs(dirname(report_infos[-1].local_full_path), exist_ok=True)
            makedirs(dirname(report_infos[-1].fserver_full_path), exist_ok=True)
    return report_infos
