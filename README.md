# expy — Excel in Python (MVP)
 
**Goal:** Compose spreadsheets with pure Python—no cell coordinates. You assemble rows/columns/cells; expy handles layout. Everything is type-hinted, data-only (no formulas), and fully stylable with Tailwind-like utilities or themes.
 
## Core ideas
 
* **Positionless composition:** You never place `A1`, `B2`, etc. Components flow in order.
* **Small primitives:** `row`, `col`, `cell` are the building blocks.
* **Reusable components:** Start with `table`; more to come.
* **Styling-first:** Utility styles (Tailwind-ish) + global theme. Every primitive takes `style=[...]`.
* **Flexible layouts:** Vertical / Horizontal stacking of components.
 
## Primitives
 
```python
import expy as ex
 
 
report = (
    ex.workbook("Sales")[
        ex.sheet("Summary")[
            ex.row(style=[ex.text_xl, ex.bold])["Q3 Sales Overview"],
            ex.row()[
                ex.cell(style=[ex.muted])["Region"],
                ex.cell(style=[ex.muted])["Units"],
                ex.cell(style=[ex.muted])["Price"],
            ],
            ex.row(style=[ex.bg_emphasis, ex.text_white])[ "EMEA", 1200, 19.0 ],
            ex.row()[ "APAC",  900, 21.0 ],
            ex.row()[ "AMER", 1500, 18.5 ],
        ]
    ]
)
 
report.save("report.xlsx")
```
 
### `row`, `col`, `cell`
 
```python
e.row(style=[ex.bold, ex.bg_red])[ 1, 2, 3, 4, 5 ]
e.col(style=[ex.italic])[ "a", "b", "c" ]
e.cell(style=[ex.bold, ex.bg_red])[ "hello, world!" ]
```
 
* `row[...]` accepts a sequence (mixed types ok).
* `col[...]` stacks values vertically.
* `cell[...]` wraps a single value.
* All accept `style=[...]`.
 
## Component: `table`
 
TBD
 
## Styling
 
### Utility styles (Tailwind-like)
 
```python
e.row(style=[ex.text_sm, ex.text_slate_600, ex.bg_slate_50, ex.px_2, ex.py_1])[ ... ]
e.cell(style=[ex.text_right, ex.mono])[ 42100 ]
```
 
**Example utilities** (non-exhaustive):
 
* Typography: `text_xs/_sm/_base/_lg/_xl`, `bold`, `italic`, `mono`
* Color: `text_slate_600`, `bg_emphasis`, `bg_red`, `text_white`
* Layout: `px_2`, `py_1`, `align_middle`, `text_right`
* Table: `table_bordered`, `table_banded`, `table_compact`
 
### Themes (sane defaults)
 
```python
corp = ex.theme(
    font="Inter",
    base_size=11,
    colors={"primary": "#0D6EFD", "muted": "#6C757D"},
    table=ex.table_theme(banded=True, header_bg="#F2F4F7")
)
 
e.workbook("Styled", theme=corp)[
    e.sheet("Overview")[ e.table(header=[...])[ ... ] ]
]
```
 
* A theme sets defaults; local `style=[...]` always winex.
* Swap themes to reskin without changing data/layout code.
 
## Types & ergonomics
 
* Modern Python with type hints on all public APIex.
* Pure Python stack traces; easy to debug and test.
* Deterministic rendering for stable diffs in CI.
