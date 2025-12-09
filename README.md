# Excel Charts

Python library for creating Excel charts using xlsxwriter with an OOP approach.

## Installation

```bash
pip install -e .
```

## Quick Start

```python
import pandas as pd
import xlsxwriter
from excel_charts import LineChartWrapper, MoneyAxis

data = pd.DataFrame({
    "Month": ["Jan", "Feb", "Mar"],
    "Revenue": [10000, 23000, 15000],
})

wb = xlsxwriter.Workbook("output.xlsx")

line = LineChartWrapper(
    data=data,
    title="Monthly Revenue",
    chart_position="E1",
    money_axis=MoneyAxis.Y,
    y_axis_title="Revenue ($)",
    x_axis_title="Month"
)
line.add_to_workbook(wb)

wb.close()
```

## Chart Types

### LineChartWrapper
```python
from excel_charts import LineChartWrapper, MoneyAxis

line = LineChartWrapper(
    data=data,
    title="Sales Trend",
    chart_position="E1",
    money_axis=MoneyAxis.Y,
    y_axis_title="Sales ($)",
    x_axis_title="Quarter",
    max_x_axis_value="Q3",
    x_major_unit=1,
    y_major_unit=5000
)
```

### ScatterChartWrapper
```python
from excel_charts import ScatterChartWrapper

scatter = ScatterChartWrapper(
    data=data,
    title="Price vs Quantity",
    chart_position="E1",
    x_axis_title="Quantity",
    y_axis_title="Price ($)"
)
```

### BarChartWrapper
```python
from excel_charts import BarChartWrapper, BarOrientation

bar = BarChartWrapper(
    data=data,
    title="Sales by Region",
    chart_position="E1",
    orientation=BarOrientation.VERTICAL,
    x_axis_title="Region",
    y_axis_title="Sales ($)"
)
```

### DonutChartWrapper
```python
from excel_charts import DonutChartWrapper

donut = DonutChartWrapper(
    data=data,
    title="Market Share",
    chart_position="E1"
)
```

## Advanced: Source Object

```python
from excel_charts import Source, LineChartWrapper

source = Source(data=data, sheet="Data", position="B5")

line = LineChartWrapper(
    data=source,
    title="Custom Chart",
    chart_position="E1"
)
```

## API Reference

### BaseChart Parameters
- `data`: pd.DataFrame or Source object
- `title`: Chart title
- `chart_position`: Cell where chart appears (default: "E1")
- `sheet`: Sheet name (defaults to title)
- `money_axis`: MoneyAxis.X or MoneyAxis.Y
- `x_axis_title`: X-axis label
- `y_axis_title`: Y-axis label

### LineChartWrapper Additional Parameters
- `max_x_axis_value`: Limit x-axis range
- `x_major_unit`: X-axis major tick interval
- `x_minor_unit`: X-axis minor tick interval
- `y_major_unit`: Y-axis major tick interval
- `y_minor_unit`: Y-axis minor tick interval

### BarChartWrapper Additional Parameters
- `orientation`: BarOrientation.VERTICAL or BarOrientation.HORIZONTAL

## Examples

See `examples/example.py` for complete examples.

## License

MIT
