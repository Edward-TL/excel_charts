"""line.py

Implementation of a line chart using xlsxwriter.
"""
from __future__ import annotations
from dataclasses import dataclass
from typing import Any

from excel_charts.core import BaseChart, MoneyAxis
from xlsxwriter.chart import Chart



@dataclass
class Line(BaseChart):
    """Line chart wrapper.

    Expects the DataFrame to have at least two columns: the first column for the
    X‑axis values and the second column for the Y‑axis values.
    """
    chart: Chart | None = None
    width: int | None = None
    height: int | None = None
    
    def __post_init__(self) -> None:
        super().__post_init__()

    def create_from_table(self) -> None:
        if not self.source.is_excel_table:
             msg = "Source is not an Excel table."
             msg += "When adding to worksheet, use the as_table=True option."
             raise ValueError(msg)
        
        self._create_chart()

    def _create_chart(self) -> None:
        # Create chart object
        self.chart = self.wb.add_chart({"type": "line"})
        self.chart.set_title({"name": self.title})
        
        # Configure X and Y axis
        self.set_y_axis()
        self.set_x_axis()

        # Create ranges using source helpers
        
        # Check if we have multiple series (cols > 2)
        # Col 0 is categories (X axis)
        # Col 1..N are values (Series 1..N)
        reference_cols = {
                c: col for c, col in enumerate(self.source.data.columns)
            }

        # Categories are always column 0
        categories_ref = self.source.get_category_ref(0)

        for col_idx in range(1, len(reference_cols)):

            if self.skip is not None:
                if reference_cols[col_idx] in self.skip:
                    continue
            
            # print(f"Adding: {reference_cols[col_idx]=}")
            values_ref = self.source.get_ref(col_idx)
            
            # Series name comes from the header of that column
            series_name = [
                self.source.worksheet,
                self.source.start_row - 1,
                self.source.start_col + col_idx
            ]

            self.chart.add_series({
                "name": series_name,
                "categories": categories_ref,
                "values": values_ref,
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