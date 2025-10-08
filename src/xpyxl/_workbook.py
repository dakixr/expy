from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook as _OpenpyxlWorkbook

from .nodes import WorkbookNode
from .render import render_sheet

__all__ = ["Workbook"]


class Workbook:
    """Immutable workbook aggregate with a `.save()` convenience."""

    def __init__(self, node: WorkbookNode) -> None:
        self._node = node

    @property
    def name(self) -> str:
        return self._node.name

    def save(self, path: str | Path) -> None:
        workbook = self.to_openpyxl()
        workbook.save(str(Path(path)))

    def to_openpyxl(self):  # -> openpyxl.workbook.Workbook
        workbook = _OpenpyxlWorkbook()
        default_sheet = workbook.active
        if default_sheet is not None:
            workbook.remove(default_sheet)
        for sheet in self._node.sheets:
            ws = workbook.create_sheet(title=sheet.name)
            render_sheet(ws, sheet)
        return workbook
