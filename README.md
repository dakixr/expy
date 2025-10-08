# expy — Excel in Python

Compose polished spreadsheets with pure Python—no manual coordinates. You assemble rows/columns/cells; expy handles layout, rendering, and styling with utility-style classes.

## Core ideas

- **Positionless composition:** Build sheets declaratively from `row`, `col`, `cell`, `table`, `vstack`, and `hstack`.
- **Composable styling:** Tailwind-inspired utilities (typography, colors, alignment, number formats) applied via `style=[...]`.
- **Deterministic rendering:** Pure-data trees compiled into `.xlsx` files with predictable output—ideal for tests and CI diffing.

## Getting started

```python
import expy as ex

report = (
    ex.workbook("Sales")[
        ex.sheet("Summary")[
            ex.row(style=[ex.text_2xl, ex.bold, ex.text_blue])["Q3 Sales Overview"],
            ex.row(style=[ex.text_sm, ex.text_gray])["Region", "Units", "Price"],
            ex.row(style=[ex.bg_primary, ex.text_white, ex.bold])["EMEA", 1200, 19.0],
            ex.row()["APAC", 900, 21.0],
            ex.row()["AMER", 1500, 18.5],
        ]
    ]
)

report.save("report.xlsx")
```

## Primitives

```python
ex.row(style=[ex.bold, ex.bg_warning])[1, 2, 3, 4, 5]
ex.col(style=[ex.italic])["a", "b", "c"]
ex.cell(style=[ex.text_green, ex.number_precision])[42100]
```

- `row[...]` accepts any sequence (numbers, strings, dataclasses…)
- `col[...]` stacks values vertically
- `cell[...]` wraps a single scalar
- All primitives accept `style=[...]`

## Component: `table`

`ex.table(...)` renders a header + body with optional style overrides. Combine with `vstack`/`hstack` for dashboards and reports.

```python
sales_table = ex.table(
    header=["Region", "Units", "Price"],
    header_style=[ex.text_sm, ex.text_gray, ex.align_middle],
    style=[ex.table_bordered, ex.table_compact],
)[
    ["EMEA", 1200, 19.0],
    ["APAC", 900, 21.0],
    ["AMER", 1500, 18.5],
]

layout = ex.vstack(
    ex.row(style=[ex.text_xl, ex.bold])["Q3 Sales Overview"],
    ex.space(),
    ex.hstack(
        sales_table,
        ex.cell(style=[ex.text_sm, ex.text_gray])["Generated with expy"],
        gap=2,
    ),
)
```

## Utility styles (non-exhaustive)

- **Typography:** `text_xs/_sm/_base/_lg/_xl/_2xl/_3xl`, `bold`, `italic`, `mono`
- **Text colors:** `text_red`, `text_green`, `text_blue`, `text_orange`, `text_purple`, `text_black`, `text_gray`
- **Backgrounds:** `bg_red`, `bg_primary`, `bg_muted`, `bg_success`, `bg_warning`, `bg_info`
- **Layout & alignment:** `text_left`, `text_center`, `text_right`, `align_top/middle/bottom`, `wrap`
- **Tables:** `table_bordered`, `table_banded`, `table_compact`
- **Number/date formats:** `number_comma`, `number_precision`, `percent`, `currency_usd`, `currency_eur`, `date_short`, `datetime_short`, `time_short`

Mix and match utilities freely—what you see is what you get.

## Layout helpers

- `vstack(a, b, c, gap=1)` vertically stacks components with optional blank rows between them.
- `hstack(a, b, gap=1)` arranges components side by side with configurable column gaps.
- `space(rows=1, height=None)` inserts empty rows (optionally with a fixed height).

## Types & ergonomics

- Modern Python with full type hints.
- Pure Python stack traces; easy to debug, script, and test.
- Deterministic rendering for stable diffs in CI.
