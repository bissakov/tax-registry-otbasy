import os
from typing import Dict

import win32com.client as win32


class ExcelConverter:
    def __init__(self) -> None:
        self.file_formats: Dict = {
            'xlsb': 50,
            'xlsx': 51
        }

    def convert(self, src_file: str, dst_file: str, file_type: str) -> None:
        if not os.path.isfile(path=src_file):
            raise ValueError(f'{src_file} does not exist.')

        file_format: int = self.file_formats[file_type]

        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        success: bool = False
        try:
            with excel.Workbooks.Open(src_file) as worbook:
                worbook.SaveAs(dst_file, FileFormat=file_format)
                success = True
        except Exception as e:
            print(f'Error Occured: {e}')
        finally:
            if success:
                excel.Application.Quit()
            else:
                excel.Quit()
