"""donut.py

Implementation of a donut (pie) chart using xlsxwriter.
"""

from __future__ import annotations

from excel_charts.core import BaseChart, MoneyAxis


class Donut(BaseChart):
    """Donut chart wrapper.

    Uses xlsxwriter 'doughnut' chart type.
    """

    def _create_chart(self):
        # Create chart
        chart = self._workbook.add_chart({"type": "doughnut"})
        chart.set_title({"name": self.title})
        chart.set_hole_size(50)
        
        # Ranges
        # Donut Chart: 1st col categories, 2nd col values
        cat_col = 0
        val_col = 1

        cats_ref = self.source.get_category_ref(cat_col)
        vals_ref = self.source.get_ref(val_col)
        
        series_name = [self.source.worksheet.get_name(), self.source.start_row, self.source.start_col + val_col]

        chart.add_series({
            "name": series_name,
            "categories": cats_ref,
            "values": vals_ref,
        })
        
        return chart
