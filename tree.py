```python
"""tree.py

Placeholder implementation for a hierarchical (tree) chart.

OpenPyXL does not provide a native tree chart type. Implementing a true
hierarchical chart would require either a custom drawing using the
openpyxl.drawing module or delegating to an external library (e.g.,
matplotlib) and embedding the resulting image.

For the purpose of this library we expose a ``TreeChartWrapper`` that
accepts the same constructor signature as ``BaseChart`` but raises a
``NotImplementedError`` when ``add_to_workbook`` is called. Users can
extend this class with their own rendering logic.
"""

from __future__ import annotations

from .core import BaseChart


class TreeChartWrapper(BaseChart):
    """Tree chart wrapper – not implemented.

    Sub‑class this and override ``_create_chart`` with a custom drawing
    routine if you need a true tree visualization.
    """

    def _create_chart(self):  # pragma: no cover
        raise NotImplementedError(
            "Tree charts are not directly supported by openpyxl. "
            "Consider generating the chart with matplotlib or plotly, "
            "saving it as an image, and inserting the image into the sheet."
        )
```
