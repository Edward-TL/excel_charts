"""bar.py

Implementation of vertical and horizontal bar charts using xlsxwriter.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Any
from enum import Enum

from excel_charts.core import BaseChart, MoneyAxis
from xlsxwriter.chart import Chart


class BarOrientation(str, Enum):
    """Orientation of the bar chart."""

    VERTICAL = "vertical"
    HORIZONTAL = "horizontal"


@dataclass
class Bar(BaseChart):
    """Bar chart wrapper.

    Parameters
    ----------
    orientation : BarOrientation, optional
        Determines whether the bars are drawn vertically (default) or horizontally.
    """
    chart: Chart | None = None
    width: int | None = None
    height: int | None = None
    orientation: BarOrientation = BarOrientation.VERTICAL
    
    def __post_init__(self) -> None:
        super().__post_init__()

    def create_from_table(self) -> None:
        if not self.source.is_excel_table:
             msg = "Source is not an Excel table."
             msg += "When adding to worksheet, use the as_table=True option."
             raise ValueError(msg)
        
        self._create_chart()

    def _create_chart(self) -> None:
        # Set orientation
        if self.orientation == BarOrientation.HORIZONTAL:
            chart_type = "bar"
        else:
            chart_type = "column"

        # Create chart object
        self.chart = self.wb.add_chart({"type": chart_type})
        self.chart.set_title({"name": self.title})
        
        # Configure X and Y axis
        self.set_y_axis()
        self.set_x_axis()

        # Determine cols
        if self.money_axis == MoneyAxis.Y:
            cat_col = 0
            val_col = 1
        else:
            cat_col = 1
            val_col = 0
        
        # Create ranges
        cats_ref = self.source.get_category_ref(cat_col)
        vals_ref = self.source.get_ref(val_col)
        
        # Series name comes from the header of that column
        series_name = [
            self.source.worksheet,
            self.source.start_row - 1,
            self.source.start_col + val_col
        ]

        self.chart.add_series({
            "name": series_name,
            "categories": cats_ref,
            "values": vals_ref,
        })

        self.chart.set_size(
            {
                'width': self.width,
                'height': self.height
            }
        )
        self.ws.insert_chart(
            self.chart_position,
            self.chart
        )

    def set_x_axis(self) -> None:
        """Set the X axis options."""
        if self.x_axis:
            self.chart.set_x_axis(self.x_axis.to_dict())

    def set_y_axis(self) -> None:
        """Set the Y axis options."""
        if self.y_axis:
            self.chart.set_y_axis(self.y_axis.to_dict())

