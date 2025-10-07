# Implementation Plan for expy MVP

## Objectives
- Deliver the fluent builder API shown in README (`workbook`, `sheet`, `row`, `col`, `cell`, `table`).
- Support Tailwind-like style utilities plus theme defaults when composing spreadsheets.
- Render composed structures to `.xlsx` via a deterministic openpyxl backend.

## Work Breakdown
1. **Core data model**
   - Define immutable dataclasses for `Style`, `Theme`, `TableTheme`, and nodes (`CellNode`, `RowNode`, `ColumnNode`, `SheetNode`, `WorkbookNode`).
   - Provide ergonomic builder facades (`WorkbookBuilder`, etc.) that implement `__getitem__` for the fluent API.
   - Ensure builders capture `style=[...]` lists and raw payload values.
2. **Styling system**
   - Implement `Style` objects with merge semantics and helpers to resolve localized + theme styles.
   - Expose utility instances (`text_xl`, `bold`, `bg_emphasis`, etc.) mapping to openpyxl font/fill/alignment/number formats.
   - Implement `theme(...)` and `table_theme(...)` factories that bundle defaults.
3. **Rendering backend**
   - Translate composed tree into tabular grid coordinates while preserving deterministic ordering.
   - Use openpyxl to create workbook/sheets, emit cell values, and apply aggregated styles (including banded/bordered tables, header styling, alignment utilities).
   - Implement padding utilities via column width / row height heuristics.
4. **Component: `table`**
   - Provide `table(header=[...])` builder generating header + data rows with optional table theme hints.
   - Integrate table styles with global theme and local overrides.
5. **Public API & ergonomics**
   - Export primitives and utilities from `expy.__init__`.
   - Add docstring examples mirroring README scenarios and type hints throughout.
   - Include a lightweight smoke test or usage example (if feasible) under `__main__` guard or tests directory.
6. **Polish**
   - Update `pyproject.toml` dependencies (e.g., add `openpyxl`).
   - Run formatting / linting if available; manually verify sample `report.save(...)` works.
   - Document future enhancements / limitations in README if time allows.

## Risks & Mitigations
- **Styling fidelity gaps:** Document unsupported utilities and provide graceful fallbacks.
- **Theme/style merge complexity:** Keep merge order explicit (theme defaults < component defaults < local `style`).
- **Openpyxl dependency footprint:** Ensure lazy import and narrow exposure to keep cold-start quick.

