from __future__ import annotations

import abc
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, TYPE_CHECKING, Union

import pandas as pd

if TYPE_CHECKING:
    from xlsxwriter.workbook import Workbook
    from xlsxwriter.worksheet import Worksheet
    from xlsxwriter.chart import Chart

try:
    from xlsxwriter.utility import xl_cell_to_rowcol
except ImportError:
    # Fallback if xlsxwriter not fully installed or mocked test env
    def xl_cell_to_rowcol(cell_str):
        return 0, 0


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


class Source:
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
    def __init__(self, data: pd.DataFrame, sheet: str | None = None, position: str = "A1"):
        self.data = data
        self.sheet = sheet
        self.position_str = position
        
        # Internal state after adding to workbook
        self._wb: Workbook | None = None
        self._ws: Worksheet | None = None
        self.start_row: int = 0
        self.start_col: int = 0
        self.end_row: int = 0
        self.end_col: int = 0
    
    @property
    def worksheet(self) -> Worksheet:
        if self._ws is None:
             raise ValueError("Source has not been added to a workbook yet.")
        return self._ws

    def add_to_workbook(self, wb: Workbook, default_sheet_name: str) -> None:
        """Writes the data to the workbook."""
        self._wb = wb
        sheet_name = self.sheet or default_sheet_name
        
        self._ws = wb.get_worksheet_by_name(sheet_name)
        if self._ws is None:
            self._ws = wb.add_worksheet(sheet_name)
            
        # Parse position
        try:
            self.start_row, self.start_col = xl_cell_to_rowcol(self.position_str)
        except Exception:
            self.start_row, self.start_col = 0, 0
            
        # Write headers
        for col_num, value in enumerate(self.data.columns):
            self._ws.write(self.start_row, self.start_col + col_num, value)

        # Write data
        self.end_row = self.start_row
        for row_idx, row in enumerate(self.data.itertuples(index=False), start=1):
            current_row = self.start_row + row_idx
            self.end_row = current_row
            for col_idx, value in enumerate(row):
                 self._ws.write(current_row, self.start_col + col_idx, value)
        
        self.end_col = self.start_col + len(self.data.columns) - 1

    def get_ref(self, col_offset: int) -> list:
        """Returns [sheet, start_row, col, end_row, col] for a specific column offset from start."""
        col = self.start_col + col_offset
        # We skip the header row for data references usually
        data_start_row = self.start_row + 1
        return [self._ws.get_name(), data_start_row, col, self.end_row, col]

    def get_category_ref(self, col_offset: int = 0) -> list:
         # Usually categories are the specific column without header?
         # Or with header as title? xlsxwriter usually takes values separately from name.
         return self.get_ref(col_offset)


class BaseChart(abc.ABC):
    """Abstract base class for all chart types."""

    def __init__(
        self,
        data: pd.DataFrame | Source,
        title: str,
        chart_position: str = "E1",
        sheet: str | None = None, # Deprecated in favor of Source config, but kept for ease
        color_palette: ColorPalette | None = None,
        money_axis: MoneyAxis = MoneyAxis.Y,
        x_axis_title: str | None = None,
        y_axis_title: str | None = None,
    ) -> None:
        if isinstance(data, pd.DataFrame):
            # If sheet is not provided here, Source will use chart title as default sheet name later
            self.source = Source(data, sheet=sheet)
        else:
            self.source = data
        
        self.title = title
        self.chart_position = chart_position
        self.color_palette = color_palette or ColorPalette()
        self.money_axis = money_axis
        self.x_axis_title = x_axis_title
        self.y_axis_title = y_axis_title

    @property
    def _workbook(self) -> Workbook:
        # Access via source for consistency, though we pass wb to add_to_workbook
        if self.source._wb is None:
             raise ValueError("Chart has not been added to a workbook.")
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
