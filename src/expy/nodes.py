from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .styles import Style, Theme

__all__ = [
    "CellNode",
    "RowNode",
    "ColumnNode",
    "TableNode",
    "SpacerNode",
    "SheetNode",
    "WorkbookNode",
    "SheetItem",
]


@dataclass(frozen=True)
class CellNode:
    value: Any
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class RowNode:
    cells: tuple[CellNode, ...]
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class ColumnNode:
    cells: tuple[CellNode, ...]
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class TableNode:
    rows: tuple[RowNode, ...]
    styles: tuple[Style, ...] = ()
    header: RowNode | None = None


@dataclass(frozen=True)
class SpacerNode:
    rows: int = 1
    height: float | None = None


SheetItem = RowNode | ColumnNode | TableNode | SpacerNode


@dataclass(frozen=True)
class SheetNode:
    name: str
    items: tuple[SheetItem, ...]


@dataclass(frozen=True)
class WorkbookNode:
    name: str
    sheets: tuple[SheetNode, ...]
    theme: Theme
