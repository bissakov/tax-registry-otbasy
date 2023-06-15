import os

import dotenv
import requests

import logger
from bot_notification import TelegramNotifier
from data_structures import Credentials, Process

logger.setup_logger()

session = requests.Session()
dotenv.load_dotenv()
notifier = TelegramNotifier(token=os.getenv('TOKEN_LOG'), chat_id=os.getenv(f'CHAT_ID_LOG'), session=session)
credentials = Credentials(usr=os.getenv(f'COLVIR_USR'), psw=os.getenv(f'COLVIR_PSW'))
process = Process(name='COLVIR', path=os.getenv('COLVIR_PROCESS_PATH'))
