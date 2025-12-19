from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from typing import Union, Optional, Literal
from pathlib import Path
import xlsxwriter
import pandas as pd
from pandas.io.formats.style import Styler as pd_Styler

from xlsxwriter.worksheet import Worksheet


from excel_charts.workbook import Writter

try:
    from xlsxwriter.utility import xl_cell_to_rowcol, xl_range
except ImportError:
    # Fallback if xlsxwriter not fully installed or mocked test env
    def xl_cell_to_rowcol(cell_str):
        return 0, 0
        
    def xl_range(first_row, first_col, last_row, last_col):
        return ""

NUM_FORMATS = Literal[
    '$#,##0.00',
    '$ #,##0.00,," M";[Rojo]-$ #,##0.00,," M"'
]

@dataclass
class Style:
    """
    """
    main: Optional[str] = None
    by_col: Optional[dict] = None
    apply_to_index: bool = False

@dataclass
class Table:
    """Represents the data source for a chart on a worksheet.
    
    Attributes
    ----------
    data : pd.DataFrame
        The data source.
    sheet : str | None
        Name of the sheet where data will be written. If None, it will be inferred.
    position : str
        The starting cell position for the data (e.g., "A1").
    """
    name: str
    data: pd.DataFrame | pd_Styler
    wb: Writter | Workbook
    worksheet: str = "Sheet1"
    position: str = "A1"
    file: Optional[str | Path] = None
    index: Optional[str | list[str]] = None
    style: Optional[pd_Styler | dict | str | Style] = None
    ws: Worksheet = field(init=False)
    excel_name: str = field(init=False)
    # Internal state after adding to workbook
    start_row: int = field(init=False, default=0)
    start_col: int = field(init=False, default=0)
    end_row: int = field(init=False, default=0)
    end_col: int = field(init=False, default=0)
    _range: str = field(init=False, default="")
    is_excel_table: bool = field(init=False, default=False)
    
    def __post_init__(self):
        self.set_dimensions()

        if self.index is None:
            self.index = self.data.columns[0]

        if self.style is None and isinstance(self.data, pd_Styler):
            self.style = copy(self.data)
            self.data = self.data.data
        
        # print(type(self.wb))
        if isinstance(self.wb, Writter):
            self.wb = copy(self.wb.wb)
            # print(type(self.wb))
        
        self.ws = self.wb.get_worksheet_by_name(self.worksheet)
        self.excel_name = self.name.lower().replace(' ', '_')

    def add_to_worksheet(
            self,
            as_table: bool = False,
            add_title: bool = True,
            ) -> None:
        """Writes the data to the workbook."""
        # Write headers
        cols = {}
        for col_num, value in enumerate(self.data.columns):
            self.ws.write(self.start_row, self.start_col + col_num, value)
            cols[col_num] = value
        main_format = None
        col_formats = {}
        if isinstance(self.style, Style):
            if isinstance(self.style.by_col, dict):
                col_formats = {
                    col: self.wb.add_format(_format) for col, _format in self.style.by_col.items()
                }

            if isinstance(self.style.main, str):
                main_format = self.wb.add_format({'num_format': self.style.main})

        # Write data
        self.end_row = self.start_row
        for row_idx, row in enumerate(self.data.itertuples(index=False), start=1):
            current_row = self.start_row + row_idx
            self.end_row = current_row

            for col_idx, value in enumerate(row):
                if col_idx in cols:
                    col_name = cols[col_idx]
                
                # Determine format for this cell
                cell_format = main_format
                if col_name in col_formats:
                    cell_format = col_formats[col_name]

                # print(col_name, cell_format)
                self.ws.write(current_row, self.start_col + col_idx, value, cell_format)
        
        self.end_col = self.start_col + len(self.data.columns) - 1
        
        
        if as_table:
            self.create_table(self.ws)

        if add_title:
            self.add_title()
        # print(type(self.wb), type(self.ws), as_table)
        

    def get_ref(self, col_offset: int = 0) -> list | str:
        """Returns [sheet, start_row, col, end_row, col] for a specific column offset from start."""
        col = self.start_col + col_offset
        
        if self.is_excel_table:
            # col_offset 0 corresponds to the FIRST column in the dataframe
            # Assuming col_offset matches the index in self.data.columns
            if 0 <= col_offset < len(self.data.columns):
                col_name = self.data.columns[col_offset]
                # Return structured reference e.g. "table_name[column_name]"
                return f"{self.excel_name}[{col_name}]"
        
        # We skip the header row for data references usually
        data_start_row = self.start_row + 1
        return [self.worksheet, data_start_row, col, self.end_row, col]

    def get_category_ref(self, col_offset: int = 0) -> list:
         # Usually categories are the specific column without header?
         # Or with header as title? xlsxwriter usually takes values separately from name.
         return self.get_ref(col_offset)

    def add_title(self) -> None:
        """
        Adds a merged title cell above the table and shifts the table down.
        """
        # 1. Merge the cells at the top (current start_row)
        title_format = self.wb.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        self.ws.merge_range(
            self.start_row -1, self.start_col,
            self.start_row -1, self.end_col,
            self.name,
            title_format
        )
        
        # 2. Shift the table internal dimensions down by 1 row
        self.start_row += 1
        self.end_row += 1
        
        # 3. Recalculate the Excel range string for the table data
        self._range = xl_range(self.start_row, self.start_col, self.end_row, self.end_col)

    def set_dimensions(self) -> None:
        """Sets start_row, start_col, end_row, end_col and _range."""
        try:
            self.start_row, self.start_col = xl_cell_to_rowcol(self.position)
        except Exception:
            self.start_row, self.start_col = 0, 0

        # DataFrame shape
        rows, cols = self.data.shape
        
        # header is row 0 relative to start_row
        # data is rows 1..rows
        # so total rows used is rows + 1 (if headers included)
        
        if rows > 0:
            self.end_row = self.start_row + rows
        else:
            self.end_row = self.start_row

        if cols > 0:
            self.end_col = self.start_col + cols - 1
        else:
            self.end_col = self.start_col
            
        self._range = xl_range(self.start_row, self.start_col, self.end_row, self.end_col)

    def create_table(self, ws: Optional[Worksheet]=None) -> None:
        """Creates an Excel table with the data."""
        # Resolve formats for table columns so they match the cells
        main_format = None
        col_formats = {}
        # Access the workbook to add formats. 
        # Note: self._wb should be an xlsxwriter Workbook instance by now.
        
        if isinstance(self.style, Style):
            if isinstance(self.style.by_col, dict):
                col_formats = {
                    col: self.wb.add_format(_format) for col, _format in self.style.by_col.items()
                }

            if isinstance(self.style.main, str):
                main_format = self.wb.add_format({'num_format': self.style.main})

        columns = []
        for col in self.data.columns:
            col_def = {"header": str(col)}
            
            # Select format: specific column format > main format
            fmt = col_formats.get(col, main_format)
            if fmt:
                col_def["format"] = fmt
            
            columns.append(col_def)
        
        options = {
            # Do not pass "data" here. The data is already written by add_to_workbook 
            # with correct cell-level formatting. Passing "data" here would overwrite it.
            "columns": columns,
            "name": self.excel_name
        }
        
        if ws is None:
            ws = self.ws
        
        ws.add_table(self._range, options)
        self.is_excel_table = True
