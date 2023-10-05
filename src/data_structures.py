from dataclasses import dataclass
from os.path import join
from typing import Optional

from pywinauto import Application


@dataclass
class Credentials:
    user: str
    password: str


@dataclass
class Process:
    name: str
    path: str


@dataclass
class DateRange:
    from_date: str
    to_date: str


@dataclass
class ReportInfo:
    report_type: str
    branch: str
    local_full_path: str
    xlsb_full_path: str
    fserver_full_path: str
    range: DateRange
    app: Optional[Application] = None


@dataclass
class FilesInfo:
    path: str
    name: str
    full_path: str or None = None
    pid: int or None = None

    def __post_init__(self) -> None:
        self.full_path: str = join(self.path, self.name)
        pid: str = self.name[:self.name.find('_')]
        self.pid = int(pid)
