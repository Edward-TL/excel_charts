"""bar.py

Implementation of vertical and horizontal bar charts using xlsxwriter.
"""

from __future__ import annotations

from enum import Enum

from excel_charts.core import BaseChart, MoneyAxis


class BarOrientation(str, Enum):
    """Orientation of the bar chart."""

    VERTICAL = "vertical"
    HORIZONTAL = "horizontal"


class Bar(BaseChart):
    """Bar chart wrapper.

    Parameters
    ----------
    orientation : BarOrientation, optional
        Determines whether the bars are drawn vertically (default) or horizontally.
    """

    def __init__(self, *args, orientation: BarOrientation = BarOrientation.VERTICAL, **kwargs):
        super().__init__(*args, **kwargs)
        self.orientation = orientation

    def _create_chart(self):
        # Set orientation
        if self.orientation == BarOrientation.HORIZONTAL:
            chart_type = "bar"
        else:
            chart_type = "column"

        chart = self._workbook.add_chart({"type": chart_type})
        chart.set_title({"name": self.title})
        if self.y_axis_title:
             chart.set_y_axis({"name": self.y_axis_title})
        if self.x_axis_title:
             chart.set_x_axis({"name": self.x_axis_title})

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
        
        series_name = [self.source.worksheet.get_name(), self.source.start_row, self.source.start_col + val_col]

        chart.add_series({
            "name": series_name,
            "categories": cats_ref,
            "values": vals_ref,
        })
        
        return chart
