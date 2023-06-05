import os
import re
from os import unlink
from os.path import exists, getsize, join
from time import sleep
from typing import List, Optional, Tuple

import win32com.client as win32
from pywinauto import Application, Desktop
from pywinauto.application import ProcessNotFoundError
from pywinauto.controls.hwndwrapper import DialogWrapper, InvalidWindowHandle
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
from pywinauto.timings import TimeoutError as TimingsTimeoutError

from src.data_structures import FilesInfo, ReportInfo
from src.main import credentials, notifier, process
from src.utils import choose_mode, get_current_process_pid, get_window, is_correct_file, \
    is_errored, kill, kill_all_processes, kill_process


def login() -> None:
    desktop: Desktop = Desktop(backend='win32')
    try:
        login_win = desktop.window(title='Вход в систему')
        login_win.wait(wait_for='exists', timeout=20)
        login_win['Edit2'].wrapper_object().set_text(text=credentials.usr)
        login_win['Edit'].wrapper_object().set_text(text=credentials.psw)
        login_win['OK'].wrapper_object().click()
    except ElementAmbiguousError:
        windows: List[DialogWrapper] = Desktop(backend='win32').windows()
        for win in windows:
            if 'Вход в систему' not in win.window_text():
                continue
            kill_process(pid=win.process_id())
        raise ElementNotFoundError


def confirm_warning(app: Application) -> None:
    found = False
    for window in app.windows():
        if found:
            break
        if window.window_text() != 'Colvir Banking System':
            continue
        win = app.window(handle=window.handle)
        for child in win.descendants():
            if child.window_text() == 'OK':
                found = True
                win.close()
                if win.exists():
                    win.close()
                break


def open_colvir(pids: List[int], retry_count: int = 0) -> Optional[Application]:
    if retry_count == 3:
        raise RuntimeError('Не удалось запустить Colvir')

    try:
        Application(backend='win32').start(cmd_line=process.path)
        login()
        sleep(4)
    except (ElementNotFoundError, TimingsTimeoutError):
        retry_count += 1
        kill(pids=pids)
        app = open_colvir(pids=pids, retry_count=retry_count)
        return app
    try:
        pid: int = get_current_process_pid(proc_name='COLVIR', pids=pids)
        app: Application = Application(backend='win32').connect(process=pid)
        try:
            if app.Dialog.window_text() == 'Произошла ошибка':
                retry_count += 1
                kill(pids=pids)
                app = open_colvir(pids=pids, retry_count=retry_count)
                return app
        except MatchError:
            pass
    except ProcessNotFoundError:
        sleep(1)
        pid = get_current_process_pid(proc_name='COLVIR', pids=pids)
        app: Application = Application(backend='win32').connect(process=pid)
    try:
        confirm_warning(app=app)
        sleep(2)
        if is_errored(app=app):
            raise ElementNotFoundError
    except (ElementNotFoundError, MatchError):
        retry_count += 1
        kill(pids=pids)
        app = open_colvir(pids=pids, retry_count=retry_count)
    return app


def prepare_report(app: Application, report_info: ReportInfo) -> None:
    choose_mode(app=app, mode='TREPRT')

    report_win = get_window(app=app, title='Выбор отчета')
    sleep(1)
    report_win.send_keystrokes('{F9}')

    sleep(1)
    copper_filter_win = get_window(app=app, title='Фильтр')
    copper_filter_win['Edit4'].set_text(text=report_info.report_type)
    sleep(1)
    copper_filter_win['OK'].wrapper_object().click()
    sleep(1)

    report_win['Предварительный просмотр'].wrapper_object().click()
    report_win['Экспорт в файл...'].wrapper_object().click()

    if exists(report_info.report_local_full_path):
        unlink(report_info.report_local_full_path)

    file_win = get_window(app=app, title='Файл отчета ')
    file_win['Edit2'].set_text(text=report_info.report_local_folder_path)
    sleep(1)
    file_win['Edit4'].set_text(text=f'{app.process}_{report_info.report_name}')
    try:
        file_win['ComboBox'].select(11)
        sleep(1)
    except (IndexError, ValueError):
        pass
    file_win['OK'].click()

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
    params_win['OK'].wrapper_object().click()


def close_sessions(pids: List[int], done_files: List, excel: win32.Dispatch,
                   pids_number: int) -> Tuple[List[int], List]:
    _done_files = list(done_files)
    files_info: List[FilesInfo] = []
    for path, subdirs, files in os.walk(r'C:\Users\robot.ad\Desktop\tax registry\reports'):
        for name in files:
            if name in _done_files:
                continue
            files_info.append(FilesInfo(path=path, name=name))

    for file_info in files_info:
        path = file_info.path
        name = file_info.name
        full_path = file_info.full_path
        pid = file_info.pid
        if not exists(path=full_path) and getsize(filename=full_path) == 0:
            continue
        try:
            app = Application(backend='win32').connect(process=pid)
            if not any('Выбор отчета' in win.window_text() for win in app.windows()):
                continue
            try:
                os.rename(src=full_path, dst=full_path)
            except OSError:
                continue
            if not is_correct_file(root=path, xls_file_path=name, excel=excel):
                continue
            kill_process(pid=pid)
            pids.remove(pid)
            message = f'{len(_done_files) + 1}/{pids_number}\t{pid} was terminated'
            notifier.send_message(message=message)
            _done_files.append(name)
        except (ValueError, ProcessNotFoundError, InvalidWindowHandle):
            continue
    return pids, _done_files


def convert_reports():
    kill_all_processes(proc_name='EXCEL')
    excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False

    i = 0
    for path, subdirs, files in os.walk(r'C:\Users\robot.ad\Desktop\tax registry\reports'):
        for name in files:
            xls_full_path = join(path, name)
            xlsb_full_path = re.sub(r'(.+)reports(.+?)\d+_(Z_160_DEPOFNO200[_025]*?__\d\d).xls',
                                    r'\g<1>converted_reports\g<2>\g<3>.xlsb', xls_full_path)
            i += 1
            print(i)
            try:
                wb = excel.Workbooks.Open(xls_full_path)
                wb.SaveAs(xlsb_full_path, FileFormat=50)
                wb.Close()
                print('success')
            except Exception:
                continue
    kill_all_processes(proc_name='EXCEL')


def run_colvir(report_infos: List[ReportInfo]) -> None:
    notifier.send_message('Process has started')

    pids = []
    report_len = len(report_infos)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False

    for index, report_info in enumerate(report_infos, start=1):
        app = open_colvir(pids=pids)
        pids.append(app.process)

        prepare_report(app=app, report_info=report_info)
        notifier.send_message(message=f'{index}/{report_len}')

    done_files = []
    while pids:
        pids, done_files = close_sessions(pids=pids, done_files=done_files,
                                          excel=excel, pids_number=report_len)

    notifier.send_message('Converting reports')
    convert_reports()
    notifier.send_message('Reports converted')
    notifier.send_message('Process succesfully finished')
