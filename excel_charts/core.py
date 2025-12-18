
import abc
from dataclasses import dataclass, field
from enum import Enum
import pandas as pd

from xlsxwriter.workbook import Workbook
from xlsxwriter.chart import Chart

from excel_charts.table import Table

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
