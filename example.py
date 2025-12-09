```python
"""example.py

Simple demonstration of the chart library.

Run with:
    python example.py

Make sure you have the required dependencies installed:
    pip install pandas openpyxl
"""

from pathlib import Path
import pandas as pd

from eternal_lagoon import (
    LineChartWrapper,
    ScatterChartWrapper,
    BarChartWrapper,
    BarOrientation,
    DonutChartWrapper,
    TreeChartWrapper,
    MoneyAxis,
)
from openpyxl import Workbook

# Sample data
data = pd.DataFrame({
    "Category": ["A", "B", "C", "D"],
    "Value": [10, 23, 7, 15],
})

# Create a workbook
wb = Workbook()

# Line chart
line = LineChartWrapper(data=data, sheet="Line", position="A1", money_axis=MoneyAxis.Y)
line.add_to_workbook(wb)

# Scatter chart
scatter = ScatterChartWrapper(data=data, sheet="Scatter", position="A1", money_axis=MoneyAxis.Y)
scatter.add_to_workbook(wb)

# Bar chart (vertical)
bar = BarChartWrapper(data=data, sheet="Bar", position="A1", orientation=BarOrientation.VERTICAL, money_axis=MoneyAxis.Y)
bar.add_to_workbook(wb)

# Bar chart (horizontal)
bar_h = BarChartWrapper(data=data, sheet="BarH", position="A1", orientation=BarOrientation.HORIZONTAL, money_axis=MoneyAxis.Y)
bar_h.add_to_workbook(wb)

# Donut chart
donut = DonutChartWrapper(data=data, sheet="Donut", position="A1", money_axis=MoneyAxis.Y)
donut.add_to_workbook(wb)

# Tree chart – will raise NotImplementedError if used
# tree = TreeChartWrapper(data=data, sheet="Tree", position="A1")
# tree.add_to_workbook(wb)

# Save workbook
output_path = Path("demo_charts.xlsx")
wb.save(output_path)
print(f"Workbook saved to {output_path.resolve()}")
```
