```python
"""__init__.py

Expose the chart wrapper classes for convenient import.
"""

from .core import BaseChart, ColorPalette, MoneyAxis
from .line import LineChartWrapper
from .scatter import ScatterChartWrapper
from .bar import BarChartWrapper, BarOrientation
from .donut import DonutChartWrapper
from .tree import TreeChartWrapper

__all__ = [
    "BaseChart",
    "ColorPalette",
    "MoneyAxis",
    "LineChartWrapper",
    "ScatterChartWrapper",
    "BarChartWrapper",
    "BarOrientation",
    "DonutChartWrapper",
    "TreeChartWrapper",
]
```
