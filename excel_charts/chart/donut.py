"""donut.py

Implementation of a donut (pie) chart using xlsxwriter.
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Optional

from excel_charts.core import BaseChart, MoneyAxis
from xlsxwriter.chart import Chart


@dataclass
class Donut(BaseChart):
    """Donut chart wrapper.

    Uses xlsxwriter 'doughnut' chart type.
    Expects the DataFrame to have at least two columns: the first column for
    categories and the second column for values.
    """
    chart: Optional[Chart] = None
    width: Optional[int] = None
    height: Optional[int] = None
    hole_size: Optional[int] = None
    explode: Optional[list[int]] = None
    categories_col: Optional[str] = None
    values_col: Optional[str] = None
    rotation: Optional[int] = None
    colors: Optional[dict[str]|list[str]] = None
    
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
        self.chart = self.wb.add_chart({"type": "doughnut"})
        self.chart.set_title({"name": self.title})
        
        # Ranges
        # Donut Chart: 1st col categories, 2nd col values
        if self.categories_col is not None:
            cat_col = self.columns_idx.get(self.categories_col, 0)
        else:
            cat_col = 0
        if self.values_col is not None:
            val_col = self.columns_idx.get(self.values_col, 1)
        else:
            val_col = 1

        cats_ref = self.source.get_category_ref(cat_col)
        vals_ref = self.source.get_ref(val_col)
        
        # Series name comes from the header of that column
        series_name = [
            self.source.worksheet,
            self.source.start_row - 1,
            self.source.start_col + val_col
        ]

        points = []
        if self.colors:
            for cat in self.source.data.iloc[:, cat_col].unique():
                if cat not in self.colors:
                    continue
                points.append(
                    {
                        'fill': { "color": self.colors[cat]}
                    }
                )
        
        series = {
            "name": series_name,
            "categories": cats_ref,
            "values": vals_ref,
        }
        
        if points:
            series["points"] = points
        
        self.chart.add_series(series)

        self.chart.set_size(
            {
                'width': self.width,
                'height': self.height
            }
        )

        if self.hole_size:
            self.chart.set_hole_size(self.hole_size)
        if self.rotation:
            self.chart.set_rotation(self.rotation)

        # Configure X and Y axis
        self.set_y_axis()
        self.set_x_axis()
        
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
