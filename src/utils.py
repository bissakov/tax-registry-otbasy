import shutil
from contextlib import contextmanager
from os import unlink
from os.path import exists
from time import sleep
from typing import List, Optional

import openpyxl
import psutil
import win32com.client as win32
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from psutil import Process
from pywinauto import Application, ElementNotFoundError, WindowSpecification

from src.config import get_credentials, get_process_path


def get_app(title: str, backend: str = 'win32') -> Application:
    app = None
    while not app:
        try:
            app = Application(backend=backend).connect(title=title)
        except ElementNotFoundError:
            sleep(.1)
            continue
    return app


def login(app: Optional[int] = None) -> None:
    user, password = get_credentials()

    if not app:
        app = get_app(title='Вход в систему')
    login_win = app.window(title='Вход в систему')

    login_username = login_win['Edit2']
    login_password = login_win['Edit']

    login_username.set_text(text=user)
    if login_username.window_text() != user:
        login_username.set_text('')
        login_username.type_keys(user)

    login_password.set_text(text=password)
    if login_password.window_text() != password:
        login_password.set_text('')
        login_password.type_keys(password)

    login_win['OK'].send_keystrokes('{ENTER}')

    sleep(.5)
    if login_win.exists() and app.window(title='Произошла ошибка').exists():
        raise ElementNotFoundError()


def confirm(app: Optional[Application] = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=get_process_path())
    dialog = app.window(title='Colvir Banking System', found_index=0)
    timeout = 0
    while not dialog.window(best_match='OK').exists():
        if timeout >= 5.0:
            raise ElementNotFoundError()
        timeout += .1
        sleep(.1)
    if dialog.is_visible():
        dialog.send_keystrokes('~')
    else:
        raise ElementNotFoundError()


def check_interactivity(app: Optional[Application] = None) -> None:
    if not app:
        app = Application(backend='win32').connect(path=get_process_path())
    choose_mode(app=app, mode='EXTRCT')
    sleep(1)
    if (filter_win := app.window(title='Фильтр')).exists():
        filter_win.close()
    else:
        raise ElementNotFoundError()


@contextmanager
def dispatch(application: str) -> None:
    app = win32.Dispatch(application)
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        kill_all_processes(proc_name='EXCEL')


@contextmanager
def workbook_open(excel: win32.Dispatch, file_path: str) -> None:
    wb = excel.Workbooks.Open(file_path)
    try:
        yield wb
    finally:
        wb.Close()


class OfficeNotFoundError(Exception):
    pass


def kill_process(pid: Optional[int]) -> None:
    if pid is None:
        return
    proc = Process(pid)
    proc.terminate()


def kill_all_processes(proc_name: str) -> None:
    for proc in psutil.process_iter():
        if proc_name in proc.name():
            try:
                proc.terminate()
            except psutil.AccessDenied:
                continue


def get_current_process_pid(proc_name: str, pids: List[int]) -> int or None:
    return next((p.pid for p in psutil.process_iter() if proc_name in p.name() and p.pid not in pids), None)


def get_window(title: str, app: Application, wait_for: str = 'exists', timeout: int = 20,
               regex: bool = False, found_index: int = 0) -> WindowSpecification:
    window = app.window(title=title, found_index=found_index) \
        if not regex else app.window(title_re=title, found_index=found_index)
    window.wait(wait_for=wait_for, timeout=timeout)
    sleep(.5)
    return window


def choose_mode(app: Application, mode: str) -> None:
    mode_win = app.window(title='Выбор режима')
    mode_win['Edit2'].set_text(text=mode)
    mode_win['Edit2'].send_keystrokes('~')


def kill(pids: List[int]) -> None:
    pid = get_current_process_pid(proc_name='COLVIR', pids=pids)
    kill_process(pid=pid)


def is_errored(app: Application) -> bool:
    for win in app.windows():
        text = win.window_text().strip()
        if text and 'Произошла ошибка' in text:
            return True
    return False


def is_correct_file(excel_full_file_path: str, excel: win32.Dispatch) -> bool:
    extension = excel_full_file_path.split('.')[-1]
    excel_full_file_path_no_ext = '.'.join(excel_full_file_path.split('.')[0:-1])
    excel_copy_path = f'{excel_full_file_path_no_ext}_copy.{extension}'
    shutil.copyfile(src=excel_full_file_path, dst=excel_copy_path)
    xlsx_file_path = f'{excel_full_file_path_no_ext}.xlsx'

    if not exists(path=xlsx_file_path):
        wb = excel.Workbooks.Open(excel_copy_path)
        wb.SaveAs(xlsx_file_path, FileFormat=51)
        wb.Close()

    workbook: Workbook = openpyxl.load_workbook(xlsx_file_path, data_only=True)
    sheet: Worksheet = workbook.active
    unlink(excel_copy_path)
    unlink(xlsx_file_path)

    return next((True for row in sheet.iter_rows(max_row=50) for cell in row if cell.alignment.horizontal), False)
