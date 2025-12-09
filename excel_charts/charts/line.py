"""line.py

Implementation of a line chart using xlsxwriter.
"""
from __future__ import annotations

from typing import Any

from excel_charts.core import BaseChart, MoneyAxis


class Line(BaseChart):
    """Line chart wrapper.

    Expects the DataFrame to have at least two columns: the first column for the
    X‑axis values and the second column for the Y‑axis values.
    """

    def __init__(
        self, 
        *args, 
        max_x_axis_value: Any | None = None,
        x_major_unit: float | None = None,
        x_minor_unit: float | None = None,
        y_major_unit: float | None = None,
        y_minor_unit: float | None = None,
        **kwargs
    ) -> None:
        super().__init__(*args, **kwargs)
        self.max_x_axis_value = max_x_axis_value
        self.x_major_unit = x_major_unit
        self.x_minor_unit = x_minor_unit
        self.y_major_unit = y_major_unit
        self.y_minor_unit = y_minor_unit

    def _create_chart(self):
        # Create chart object
        chart = self._workbook.add_chart({"type": "line"})
        chart.set_title({"name": self.title})
        
        # Configure Y axis
        y_axis_options = {}
        if self.y_axis_title:
             y_axis_options["name"] = self.y_axis_title
        if self.y_major_unit is not None:
             y_axis_options["major_unit"] = self.y_major_unit
        if self.y_minor_unit is not None:
             y_axis_options["minor_unit"] = self.y_minor_unit
        
        if y_axis_options:
            chart.set_y_axis(y_axis_options)
        
        # Configure X axis
        x_axis_options = {}
        if self.x_axis_title:
             x_axis_options["name"] = self.x_axis_title
        if self.max_x_axis_value is not None:
             x_axis_options["max"] = self.max_x_axis_value
        if self.x_major_unit is not None:
             x_axis_options["major_unit"] = self.x_major_unit
        if self.x_minor_unit is not None:
             x_axis_options["minor_unit"] = self.x_minor_unit
        
        if x_axis_options:
            chart.set_x_axis(x_axis_options)
        
        # Determine cols
        if self.money_axis == MoneyAxis.Y:
            # X is col 0, Y is col 1
            cat_col = 0
            val_col = 1
        else:
             # X is col 1, Y is col 0
            cat_col = 1
            val_col = 0

        # Create ranges using source helpers
        categories_ref = self.source.get_category_ref(cat_col)
        values_ref = self.source.get_ref(val_col)

        # Series name comes from the header row of the value column
        series_name = [
            self.source.worksheet.get_name(),
            self.source.start_row,
            self.source.start_col + val_col
            ]

        chart.add_series({
            "name": series_name,
            "categories": categories_ref,
            "values": values_ref,
        })
        
        return chart
