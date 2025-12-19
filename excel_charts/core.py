
import abc
from dataclasses import dataclass, field
from typing import Optional
from enum import Enum
import pandas as pd
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.workbook import Workbook
from xlsxwriter.chart import Chart

from excel_charts.table import Table
from excel_charts.workbook import Writter

class MoneyAxis(str, Enum):
    """Enum to indicate which axis represents monetary values."""
    X = "x"
    Y = "y"

@dataclass
class Axis:
    name: str | None = None
    title: str | None = None
    units: str | None = None
    major_unit: float | None = None
    minor_unit: float | None = None
    is_money: bool = False

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "units": self.units,
            "major_unit": self.major_unit,
            "minor_unit": self.minor_unit
        }

@dataclass
class Units:
    x: Axis | None = None
    y: Axis | None = None

    def to_dict(self) -> dict:
        return {
            "x": self.x.to_dict() if self.x else None,
            "y": self.y.to_dict() if self.y else None
        }

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
    category_colors: dict = field(default_factory=dict)



@dataclass
class BaseChart(abc.ABC):
    """Abstract base class for all chart types."""

    source: Table
    chart_position: str = "A1"
    worksheet: str = "Sheet1"
    title: Optional[str] = None
    x_axis: Axis | None = None
    y_axis: Axis | None = None
    money_axis: MoneyAxis = MoneyAxis.Y
    color_palette: ColorPalette | None = None
    wb: Writter | Workbook = field(init=False)
    ws: Worksheet = field(init=False)
    chart: Chart | None = None
    skip: Optional[list[str]] = None
    
    def __post_init__(self) -> None:
        if self.title is None:
            self.title = self.source.name
        if self.color_palette is None:
            self.color_palette = ColorPalette()

        self.wb = self.source.wb
        self.ws = self.wb.get_worksheet_by_name(self.worksheet)

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
            self.ws.insert_chart(self.chart_position, chart)
