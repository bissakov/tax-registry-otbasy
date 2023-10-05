import logging
from time import sleep
from typing import Optional
from urllib.parse import urljoin

import httpx
import numpy
from httpx import ConnectError, ReadError, ReadTimeout, RemoteProtocolError

from src.config import get_telegram_secrets


def _send_message(message: str, message_id: Optional[int] = None, client: Optional[httpx.Client] = None) -> httpx.Response:
    token, chat_id = get_telegram_secrets()
    api_url = f'https://api.telegram.org/bot{token}/'
    send_data = {'chat_id': chat_id, 'text': message}
    if message_id:
        send_data['message_id'] = message_id
    url = urljoin(api_url, 'sendMessage' if not message_id else 'editMessageText')
    response = client.post(url, data=send_data, timeout=300) if client \
        else httpx.post(url, data=send_data, timeout=300)

    if response.status_code == 400:
        logging.info(f'response: {response}\nmessage_id: {message_id}\nmessage: {message}')
    else:
        logging.info(f'response: {response}')
    return response


def send_message(message: str, message_id: Optional[int] = None,
                 client: Optional[httpx.Client] = None, attempt: int = 1) -> httpx.Response:
    try:
        response = _send_message(message, message_id, client)
    except (ReadTimeout, ConnectError, RemoteProtocolError, ReadError) as exc:
        if attempt >= (20 if isinstance(exc, ReadTimeout) else 5):
            raise exc
        sleep(5)
        return send_message(message, message_id, client, attempt=attempt + 1)
    return response


class ProgressBar:
    def __init__(self, client: httpx.Client, description: str = 'Progress', bar_width: int = 10) -> None:
        self.client = client
        self.description = description
        self.bar_width = bar_width
        self.progress_bar_message_id = None
        self.message = '{description}: {step}/{total}\n' \
                       '{filled}{unfilled}\n{caption}'

    def update(self, step: int, total: int, caption: str = '') -> None:
        filled = int(numpy.clip((step / total) * 100, 0, 100)) // (100 // self.bar_width)
        unfilled = self.bar_width - filled

        message = self.message.format(description=self.description,
                                      step=step, total=total,
                                      filled='█' * filled, unfilled='░' * unfilled,
                                      caption=caption)

        if self.progress_bar_message_id is None:
            response = send_message(client=self.client, message=message)
            data = response.json()
            self.progress_bar_message_id = data['result']['message_id']
        else:
            send_message(client=self.client, message=message, message_id=self.progress_bar_message_id)
