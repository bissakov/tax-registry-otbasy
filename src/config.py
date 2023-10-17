import os
from datetime import date

import dotenv

dotenv.load_dotenv()


def get_today() -> date:
    return date.today()
    # return date(2023, 10, 5)


def get_local_path() -> str:
    return r'C:\Users\robot.ad\Desktop\tax-registry-otbasy\reports'


def get_fserver_path() -> str:
    today = get_today()
    year = today.year if today.month != 1 else today.year - 1
    # \\fserver\ДБУИО_Новая\ДБУИО_Общая папка\ДБУ_Информация УНУ\2023\ИПНи СН\200 по клиентам за 2023
    return fr'\\fserver\ДБУИО_Новая\ДБУИО_Общая папка\ДБУ_Информация УНУ\{year}\ИПНи СН\200 по клиентам за {year}'
    # return r'C:\Users\robot.ad\Desktop\tax-registry-otbasy\finished'


def get_telegram_secrets() -> tuple[str, str]:
    return os.getenv('TOKEN_LOG'), os.getenv('CHAT_ID_LOG')


def get_credentials() -> tuple[str, str]:
    return os.getenv('COLVIR_USR'), os.getenv('COLVIR_PSW')


def get_process_path() -> str:
    return os.getenv('COLVIR_PROCESS_PATH')
