"""scatter.py

Implementation of a scatter chart using xlsxwriter.
"""

from __future__ import annotations

from excel_charts.core import BaseChart, MoneyAxis


class Scatter(BaseChart):
    """Scatter chart wrapper.

    The DataFrame should contain two numeric columns. The ``money_axis``
    determines which column is considered the monetary value.
    """

    def _create_chart(self):
        # Create chart object
        chart = self._workbook.add_chart({"type": "scatter"})
        chart.set_title({"name": self.title})
        if self.x_axis_title:
             chart.set_x_axis({"name": self.x_axis_title})
        if self.y_axis_title:
             chart.set_y_axis({"name": self.y_axis_title})

        # Determine column mapping based on money_axis
        if self.money_axis == MoneyAxis.Y:
            # X is col 0, Y is col 1
            x_col = 0
            y_col = 1
        else:
            # X is col 1, Y is col 0
            x_col = 1
            y_col = 0

        # Create ranges using source helpers
        x_ref = self.source.get_category_ref(x_col)
        y_ref = self.source.get_ref(y_col)
        
        series_name = [self.source.worksheet.get_name(), self.source.start_row, self.source.start_col + y_col]

        chart.add_series({
            "name": series_name,
            "categories": x_ref,
            "values": y_ref,
        })
        
        return chart
