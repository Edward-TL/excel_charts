from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional, List
import xlsxwriter
from xlsxwriter.workbook import Workbook as XlsxWorkbook


@dataclass
class Writter:
    """
    Represents an Excel workbook using XlsxWriter.
    
    Attributes
    ----------
    file : str
        The file path where the workbook will be saved.
    writer : xlsxwriter.Workbook
        The XlsxWriter Workbook instance.
    """
    file: str
    wb: XlsxWorkbook = field(init=False)
    sheet_names: list[str] = field(default_factory=lambda: ['Sheet1'])

    def __post_init__(self):
        self.wb = xlsxwriter.Workbook(self.file)

        for sheet_name in self.sheet_names:
            self.wb.add_worksheet(sheet_name)
            print(f"Adding {sheet_name=}")

    def close(self) -> None:
        """
        Saves and closes the workbook.
        """
        self.wb.close()
