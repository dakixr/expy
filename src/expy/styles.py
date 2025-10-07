from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from typing import Any, Literal, cast

BorderStyleLiteral = Literal[
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
]
BorderStyleName = BorderStyleLiteral | Literal["none"]

BORDER_STYLE_VALUES: set[str] = {
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
    "none",
}


__all__ = [
    "Style",
    "Theme",
    "TableTheme",
    "table_theme",
    "theme",
    "combine_styles",
    "normalize_hex",
    "to_argb",
    "text_xs",
    "text_sm",
    "text_base",
    "text_lg",
    "text_xl",
    "bold",
    "italic",
    "mono",
    "muted",
    "text_slate_600",
    "text_white",
    "bg_slate_50",
    "bg_emphasis",
    "bg_red",
    "text_right",
    "align_middle",
    "table_bordered",
    "table_banded",
    "table_compact",
    "BorderStyleName",
    "BorderStyleLiteral",
]


def normalize_hex(value: str) -> str:
    text = value.strip()
    if not text:
        raise ValueError("Color values cannot be empty")
    if text.startswith("#"):
        text = text[1:]
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    if len(text) != 6:
        raise ValueError(f"Expected 6 hex characters, got '{value}'")
    return "#" + text.upper()


def to_argb(value: str) -> str:
    rgb = normalize_hex(value)[1:]
    return "FF" + rgb


DEFAULT_COLORS: dict[str, str] = {
    "primary": normalize_hex("#0D6EFD"),
    "muted": normalize_hex("#6C757D"),
    "emphasis": normalize_hex("#1F2937"),
}


def _coerce_border_style(value: str) -> BorderStyleName:
    normalized = value.strip()
    if normalized not in BORDER_STYLE_VALUES:
        options = ", ".join(sorted(BORDER_STYLE_VALUES))
        msg = f"Unsupported border style '{normalized}'. Expected one of: {options}"
        raise ValueError(msg)
    return cast(BorderStyleName, normalized)


@dataclass(frozen=True)
class TableTheme:
    """High-level defaults for tables."""

    banded: bool = True
    stripe_color: str | None = normalize_hex("#F8FAFC")
    header_bg: str | None = normalize_hex("#F2F4F7")
    header_text_color: str | None = normalize_hex("#0F172A")
    border_color: str = normalize_hex("#D0D5DD")
    border_style: BorderStyleName = "thin"
    compact_row_height: float | None = 18.0


@dataclass(frozen=True)
class Theme:
    """Workbook-wide defaults."""

    font: str = "Calibri"
    base_size: float = 11.0
    colors: Mapping[str, str] = field(default_factory=lambda: dict(DEFAULT_COLORS))
    table: TableTheme = field(default_factory=TableTheme)
    mono_font: str | None = "Consolas"
    default_text_color: str = normalize_hex("#111827")


def table_theme(
    *,
    banded: bool | None = None,
    stripe_color: str | None = None,
    header_bg: str | None = None,
    header_text_color: str | None = None,
    border_color: str | None = None,
    border_style: BorderStyleName | None = None,
    compact_row_height: float | None = None,
) -> TableTheme:
    base = TableTheme()
    return TableTheme(
        banded=base.banded if banded is None else banded,
        stripe_color=normalize_hex(stripe_color) if stripe_color is not None else base.stripe_color,
        header_bg=normalize_hex(header_bg) if header_bg is not None else base.header_bg,
        header_text_color=normalize_hex(header_text_color)
        if header_text_color is not None
        else base.header_text_color,
        border_color=normalize_hex(border_color) if border_color is not None else base.border_color,
        border_style=_coerce_border_style(border_style) if border_style is not None else base.border_style,
        compact_row_height=compact_row_height if compact_row_height is not None else base.compact_row_height,
    )


def theme(
    *,
    font: str = "Calibri",
    base_size: float = 11.0,
    colors: Mapping[str, str] | None = None,
    table: TableTheme | None = None,
    mono_font: str | None = "Consolas",
    default_text_color: str = "#111827",
) -> Theme:
    palette = dict(DEFAULT_COLORS)
    if colors:
        for key, value in colors.items():
            palette[key] = normalize_hex(value)
    return Theme(
        font=font,
        base_size=base_size,
        colors=palette,
        table=table or TableTheme(),
        mono_font=mono_font,
        default_text_color=normalize_hex(default_text_color),
    )


@dataclass(frozen=True)
class Style:
    """Composable, Tailwind-like style primitive."""

    name: str = ""
    font_name: str | None = None
    font_size: float | None = None
    font_size_delta: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    mono: bool | None = None
    text_color: str | None = None
    text_color_key: str | None = None
    fill_color: str | None = None
    fill_color_key: str | None = None
    horizontal_align: str | None = None
    vertical_align: str | None = None
    indent: int | None = None
    wrap_text: bool | None = None
    number_format: str | None = None
    border: BorderStyleName | None = None
    border_color: str | None = None
    table_banded: bool | None = None
    table_bordered: bool | None = None
    table_compact: bool | None = None

    def merge(self, other: Style) -> Style:
        """Combine two styles, where ``other`` wins on conflicts."""

        base_delta = 0.0 if self.font_size_delta is None else self.font_size_delta
        other_delta = 0.0 if other.font_size_delta is None else other.font_size_delta
        merged_delta = base_delta + other_delta
        delta_value = merged_delta if (self.font_size_delta is not None or other.font_size_delta is not None) else None
        return Style(
            name=other.name or self.name,
            font_name=other.font_name or self.font_name,
            font_size=other.font_size if other.font_size is not None else self.font_size,
            font_size_delta=delta_value,
            bold=other.bold if other.bold is not None else self.bold,
            italic=other.italic if other.italic is not None else self.italic,
            mono=other.mono if other.mono is not None else self.mono,
            text_color=other.text_color or self.text_color,
            text_color_key=other.text_color_key or self.text_color_key,
            fill_color=other.fill_color or self.fill_color,
            fill_color_key=other.fill_color_key or self.fill_color_key,
            horizontal_align=other.horizontal_align or self.horizontal_align,
            vertical_align=other.vertical_align or self.vertical_align,
            indent=other.indent if other.indent is not None else self.indent,
            wrap_text=other.wrap_text if other.wrap_text is not None else self.wrap_text,
            number_format=other.number_format or self.number_format,
            border=other.border if other.border is not None else self.border,
            border_color=other.border_color if other.border_color is not None else self.border_color,
            table_banded=other.table_banded if other.table_banded is not None else self.table_banded,
            table_bordered=other.table_bordered if other.table_bordered is not None else self.table_bordered,
            table_compact=other.table_compact if other.table_compact is not None else self.table_compact,
        )


def combine_styles(styles: Iterable[Style], *, base: Style | None = None) -> Style:
    combined = base or Style()
    for style in styles:
        combined = combined.merge(style)
    return combined


def _style(name: str, **kwargs: Any) -> Style:
    return Style(name=name, **kwargs)


text_xs = _style("text_xs", font_size_delta=-2.0)
text_sm = _style("text_sm", font_size_delta=-1.0)
text_base = _style("text_base", font_size_delta=0.0)
text_lg = _style("text_lg", font_size_delta=1.0)
text_xl = _style("text_xl", font_size_delta=4.0)

bold = _style("bold", bold=True)
italic = _style("italic", italic=True)
mono = _style("mono", mono=True)

muted = _style("muted", text_color_key="muted")
text_slate_600 = _style("text_slate_600", text_color=normalize_hex("#475569"))
text_white = _style("text_white", text_color=normalize_hex("#FFFFFF"))

bg_slate_50 = _style("bg_slate_50", fill_color=normalize_hex("#F8FAFC"))
bg_emphasis = _style("bg_emphasis", fill_color_key="primary", text_color=normalize_hex("#FFFFFF"))
bg_red = _style("bg_red", fill_color=normalize_hex("#F04438"), text_color=normalize_hex("#FFFFFF"))

text_right = _style("text_right", horizontal_align="right")
align_middle = _style("align_middle", vertical_align="center")

table_bordered = _style("table_bordered", table_bordered=True)
table_banded = _style("table_banded", table_banded=True)
table_compact = _style("table_compact", table_compact=True)
