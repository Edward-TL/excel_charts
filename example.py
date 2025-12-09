"""example.py

Simple demonstration of the chart library.

Run with:
    python example.py

Make sure you have the required dependencies installed:
    pip install pandas xlsxwriter
"""

from pathlib import Path
import pandas as pd
import xlsxwriter

from eternal_lagoon import (
    LineChartWrapper,
    ScatterChartWrapper,
    BarChartWrapper,
    BarOrientation,
    DonutChartWrapper,
    TreeChartWrapper,
    MoneyAxis,
)

# Sample data
data = pd.DataFrame({
    "Category": ["A", "B", "C", "D"],
    "Value": [10, 23, 7, 15],
})

# Create a workbook
output_path = Path("demo_charts.xlsx")
wb = xlsxwriter.Workbook(output_path)

# Line chart
line = LineChartWrapper(
    data=data,
    title="Line Chart",
    chart_position="E1",
    money_axis=MoneyAxis.Y,
    y_axis_title="Y Axis",
    x_axis_title="X Axis",
    max_x_axis_value="C"
)
line.add_to_workbook(wb)

# Scatter chart
scatter = ScatterChartWrapper(
    data=data,
    # Implicitly uses chart title as sheet name if sheet not specified
    title="Scatter Chart",
    chart_position="E1",
    money_axis=MoneyAxis.Y,
    y_axis_title="Y Axis",
    x_axis_title="X Axis"
)
scatter.add_to_workbook(wb)

# Bar chart (vertical)
bar = BarChartWrapper(
    data=data,
    title="Bar Chart", 
    chart_position="E1", 
    orientation=BarOrientation.VERTICAL, 
    money_axis=MoneyAxis.Y,
    x_axis_title="Categories",
    y_axis_title="Values"
)
bar.add_to_workbook(wb)

# Bar chart (horizontal)
bar_h = BarChartWrapper(
    data=data, 
    title="Horizontal Bar Chart",
    chart_position="E1", 
    orientation=BarOrientation.HORIZONTAL, 
    money_axis=MoneyAxis.Y,
    x_axis_title="Values",
    y_axis_title="Categories"
)
bar_h.add_to_workbook(wb)

# Donut chart
donut = DonutChartWrapper(
    data=data, 
    title="Donut Chart",
    chart_position="E1", 
    money_axis=MoneyAxis.Y
)
donut.add_to_workbook(wb)

# Save workbook
wb.close()
print(f"Workbook saved to {output_path.resolve()}")
