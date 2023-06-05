from dataclasses import dataclass
from os.path import join


@dataclass
class Credentials:
    usr: str
    psw: str


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
    report_name: str
    report_local_folder_path: str
    report_local_full_path: str
    report_fserver_folder_path: str
    report_fserver_full_path: str
    range: DateRange


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
