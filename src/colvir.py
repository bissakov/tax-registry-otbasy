import logging
import os
import shutil
from os import unlink
from os.path import basename, dirname, exists
from time import sleep
from typing import List

import httpx
import win32com.client as win32
from pywinauto import Application
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
from pywinauto.timings import TimeoutError as TimingsTimeoutError

from src.bot_notification import ProgressBar
from src.data_structures import ReportInfo
from src.config import get_process_path
from src.utils import check_interactivity, choose_mode, confirm, dispatch, get_window, is_correct_file, login, workbook_open


def prepare_report(app: Application, report_info: ReportInfo) -> None:
    choose_mode(app=app, mode='TREPRT')

    report_win = get_window(app=app, title='Выбор отчета')
    report_win.send_keystrokes('{F9}')

    copper_filter_win = get_window(app=app, title='Фильтр')
    copper_filter_win['Edit4'].set_text(text=report_info.report_type)
    copper_filter_win['OK'].send_keystrokes('{ENTER}')

    while True:
        report_win['Предварительный просмотр'].click()
        sleep(.1)
        if report_win['Экспорт в файл...'].is_enabled():
            break

    report_win['Экспорт в файл...'].send_keystrokes('{ENTER}')

    if exists(report_info.local_full_path):
        unlink(report_info.local_full_path)

    file_win = get_window(app=app, title='Файл отчета ')
    file_win['Edit2'].set_text(text=dirname(report_info.local_full_path))
    sleep(1)
    file_win['Edit4'].set_text(text=basename(report_info.local_full_path))
    try:
        file_win['ComboBox'].select(11)
        sleep(1)
    except (IndexError, ValueError):
        pass
    file_win['OK'].send_keystrokes('{ENTER}')

    try:
        params_win = get_window(app=app, title='Параметры отчета ')
    except TimingsTimeoutError:
        file_win['OK'].send_keystrokes('{ENTER}')
        params_win = get_window(app=app, title='Параметры отчета ')

    if report_info.report_type == 'Z_160_DEPOFNO200':
        branch_input_box = 'Edit6'
        from_date_input_box = 'Edit2'
        to_date_input_box = 'Edit4'
    else:
        branch_input_box = 'Edit2'
        from_date_input_box = 'Edit4'
        to_date_input_box = 'Edit6'

    params_win[branch_input_box].set_text(text=report_info.branch)
    params_win[from_date_input_box].set_text(text=report_info.range.from_date)
    params_win[to_date_input_box].set_text(text=report_info.range.to_date)
    sleep(1)
    params_win['OK'].send_keystrokes('{ENTER}')


def export_tax_registry(report_info: ReportInfo) -> Application:
    retry_count: int = 0
    app = None
    while retry_count < 5:
        try:
            app = Application().start(cmd_line=get_process_path())
            login(app)
            confirm(app)
            check_interactivity(app)
            prepare_report(app=app, report_info=report_info)
            break
        except (ElementNotFoundError, MatchError, ElementAmbiguousError):
            retry_count += 1
            if app:
                app.kill()
            continue
    if retry_count == 10:
        raise Exception('max_retries exceeded')
    return app


def is_file_exported(full_file_name: str, excel: win32.CDispatch) -> bool:
    if not os.path.exists(path=full_file_name):
        return False
    if os.path.getsize(filename=full_file_name) == 0:
        return False
    try:
        os.rename(src=full_file_name, dst=full_file_name)
    except OSError:
        return False
    if not is_correct_file(excel_full_file_path=full_file_name, excel=excel):
        return False
    return True


def close_sessions(client: httpx.Client, report_infos: List[ReportInfo]) -> None:
    close_session_pbar = ProgressBar(client=client, description='Закрытие сессий')
    index = 1

    with dispatch(application='Excel.Application') as excel:
        while any(isinstance(r.app, Application) for r in report_infos):
            for i, report_info in enumerate(report_infos):
                if report_info.app is None:
                    continue

                if is_file_exported(full_file_name=report_info.local_full_path, excel=excel):
                    report_info.app.kill()
                    report_info.app = None
                    logging.info(f'{index}/{len(report_infos)} was exported')
                    close_session_pbar.update(step=index, total=len(report_infos))
                    index += 1

    # _report_infos = report_infos[:]
    # report_infos_size = len(report_infos)
    # index = 1
    #
    # with dispatch(application='Excel.Application') as excel:
    #     while _report_infos:
    #         for i, report_info in enumerate(report_infos):
    #             if report_info not in _report_infos:
    #                 continue
    #
    #             if is_file_exported(full_file_name=report_info.local_full_path, excel=excel):
    #                 _report_infos.remove(report_info)
    #                 apps[i].kill()
    #                 logging.info(f'{index}/{report_infos_size} was exported')
    #                 close_session_pbar.update(step=index, total=report_infos_size)
    #                 index += 1


def convert_reports(client: httpx.Client, report_infos: List[ReportInfo]) -> None:
    convert_pbar = ProgressBar(client=client, description='Конвертация отчетов в .xlsb')

    with dispatch(application='Excel.Application') as excel:
        for index, report_info in enumerate(report_infos, start=1):
            try:
                with workbook_open(excel=excel, file_path=report_info.local_full_path) as wb:
                    wb.SaveAs(report_info.xlsb_full_path, FileFormat=50)
                logging.info(f'{index} was converted')
                os.remove(report_info.local_full_path)
                convert_pbar.update(step=index, total=len(report_infos))
            except Exception as e:
                raise e


def transfer_files(client: httpx.Client, report_infos: List[ReportInfo]) -> None:
    transfer_pbar = ProgressBar(client=client, description='Перенос отчетов на сервер')

    for index, report_info in enumerate(report_infos, start=1):
        if exists(report_info.fserver_full_path):
            unlink(report_info.fserver_full_path)
        shutil.copyfile(src=report_info.xlsb_full_path, dst=report_info.fserver_full_path)
        transfer_pbar.update(step=index, total=len(report_infos))


def check_completion(client: httpx.Client, report_infos: List[ReportInfo]) -> None:
    check_completion_pbar = ProgressBar(client=client, description='Проверка завершения процесса')

    for index, report_info in enumerate(report_infos, start=1):
        if not exists(report_info.fserver_full_path):
            raise Exception(f'File {report_info.fserver_full_path} was not exported')
        check_completion_pbar.update(step=index, total=len(report_infos))


def run_colvir(report_infos: List[ReportInfo]) -> None:
    with httpx.Client() as client:
        colvir_pbar = ProgressBar(client=client, description='Выгрузка регистров')

        for index, report_info in enumerate(report_infos, start=1):
            app = export_tax_registry(report_info=report_info)
            report_info.app = app
            colvir_pbar.update(step=index, total=len(report_infos))

        close_sessions(client=client, report_infos=report_infos)
        convert_reports(client=client, report_infos=report_infos)
        transfer_files(client=client, report_infos=report_infos)
        check_completion(client=client, report_infos=report_infos)
