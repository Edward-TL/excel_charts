from __future__ import annotations

import abc
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, TYPE_CHECKING, Union, Optional
from pathlib import Path

import pandas as pd

if TYPE_CHECKING:
    from xlsxwriter.workbook import Workbook
    from xlsxwriter.worksheet import Worksheet
    from xlsxwriter.chart import Chart

try:
    from xlsxwriter.utility import xl_cell_to_rowcol, xl_range
except ImportError:
    # Fallback if xlsxwriter not fully installed or mocked test env
    def xl_cell_to_rowcol(cell_str):
        return 0, 0
        
    def xl_range(first_row, first_col, last_row, last_col):
        return ""


class MoneyAxis(str, Enum):
    """Enum to indicate which axis represents monetary values."""

    X = "x"
    Y = "y"

@dataclass
class ColorPalette:
    """Simple color palette definition for charts.

    Attributes
    ----------
    primary: str
        Primary color in hex (e.g., "#4A90E2").
    secondary: str
        Secondary color in hex.
    accent: str
        Accent color in hex.
    """

    primary: str = "#4A90E2"
    secondary: str = "#50E3C2"
    accent: str = "#F5A623"


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
    data: pd.DataFrame
    sheet: str | None = None
    position: str = "A1"
    _wb: Union[Workbook, str, None] = field(default=None, repr=False)
    _ws: Union[Worksheet, str, None] = field(default=None, repr=False)
    
    # Internal state after adding to workbook
    start_row: int = field(init=False, default=0)
    start_col: int = field(init=False, default=0)
    end_row: int = field(init=False, default=0)
    end_col: int = field(init=False, default=0)
    _range: str = field(init=False, default="")
    
    def __post_init__(self):
        if isinstance(self._wb, str):
            import xlsxwriter
            self._wb = xlsxwriter.Workbook(self._wb)
            
        if all(
            [
                isinstance(self._ws, str), # If it's a string, it's a sheet name
                self._wb is not None, # If it's not None, it's a workbook
                not isinstance(self._wb, str) # It's already a workbook
            ]):
            # Just keeping it more readable
            name = self._ws
            self._ws = self._wb.get_worksheet_by_name(name)
            if self._ws is None:
                self._ws = self._wb.add_worksheet(name)
        
        self.set_dimensions()
    @property
    def file(self) -> str | None:
        if self._wb is not None and hasattr(self._wb, 'filename'):
             return self._wb.filename
        return None

    @property
    def worksheet(self) -> Worksheet:
        if self._ws is None:
             raise ValueError("Source has not been added to a workbook yet.")
        return self._ws

    def add_to_workbook(self, wb: Workbook, default_sheet_name: str) -> None:
        """Writes the data to the workbook."""
        if self._wb is None:
            self._wb = wb
        
        sheet_name = self.sheet or default_sheet_name
        
        # If _ws is not set, set it now
        if self._ws is None:
             # Ensure we have a valid workbook object before calling methods
             if self._wb is not None:
                self._ws = self._wb.get_worksheet_by_name(sheet_name)
                if self._ws is None:
                    self._ws = self._wb.add_worksheet(sheet_name)
            
        # Parse position
        try:
            self.start_row, self.start_col = xl_cell_to_rowcol(self.position)
        except Exception:
            self.start_row, self.start_col = 0, 0
            
        # Write headers
        for col_num, value in enumerate(self.data.columns):
            self.worksheet.write(self.start_row, self.start_col + col_num, value)

        # Write data
        self.end_row = self.start_row
        for row_idx, row in enumerate(self.data.itertuples(index=False), start=1):
            current_row = self.start_row + row_idx
            self.end_row = current_row
            for col_idx, value in enumerate(row):
                 self.worksheet.write(current_row, self.start_col + col_idx, value)
        
        self.end_col = self.start_col + len(self.data.columns) - 1

    def get_ref(self, col_offset: int) -> list:
        """Returns [sheet, start_row, col, end_row, col] for a specific column offset from start."""
        col = self.start_col + col_offset
        # We skip the header row for data references usually
        data_start_row = self.start_row + 1
        return [self.worksheet.get_name(), data_start_row, col, self.end_row, col]

    def get_category_ref(self, col_offset: int = 0) -> list:
         # Usually categories are the specific column without header?
         # Or with header as title? xlsxwriter usually takes values separately from name.
         return self.get_ref(col_offset)

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

    def create_table(self) -> None:
        """Creates an Excel table with the data."""
        columns = [{"header": str(col)} for col in self.data.columns]
        
        # Convert data to list of lists.
        data = self.data.values.tolist()
        
        options = {
            "data": data,
            "columns": columns,
            "name": self.name
        }
        
        self.worksheet.add_table(self._range, options)


class BaseChart(abc.ABC):
    """Abstract base class for all chart types."""

    def __init__(
        self,
        source: pd.DataFrame | Table,
        title: str,
        chart_position: str = "E1",
        color_palette: ColorPalette | None = None,
        money_axis: MoneyAxis = MoneyAxis.Y,
        x_axis_title: str | None = None,
        y_axis_title: str | None = None,
    ) -> None:
        self.source = source
        self.title = title
        self.chart_position = chart_position
        self.color_palette = color_palette or ColorPalette()
        self.money_axis = money_axis
        self.x_axis_title = x_axis_title
        self.y_axis_title = y_axis_title

    @property
    def _workbook(self) -> Workbook:
        return self.source._wb

    @abc.abstractmethod
    def _create_chart(self) -> Chart:
        """Create and configure the specific xlsxwriter chart instance."""
        pass

    def add_to_workbook(self, wb: Workbook) -> None:
        """Add the chart to the supplied workbook."""
        # Initialize source (writes data)
        # Use chart title as default sheet name if source doesn't have one
        self.source.add_to_workbook(wb, default_sheet_name=self.title)
        
        # Create chart
        chart = self._create_chart()
        
        # Insert chart
        if chart:
             self.source.worksheet.insert_chart(self.chart_position, chart)
