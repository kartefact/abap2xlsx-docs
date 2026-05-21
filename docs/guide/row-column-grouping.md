# Row and Column Grouping (Outlines)

Excel's **grouping** feature lets you collapse and expand sections of rows or columns using the outline controls on the sheet margin. abap2xlsx supports this through `zcl_excel_row` and `zcl_excel_column` objects (one per row/column) and through high-level outline methods on `zcl_excel_worksheet`.

## Two Approaches

| Approach | When to use |
|---|---|
| **High-level** — `add_row_outline` / `add_column_outline` on the worksheet | Define a named, collapsible group over a range of rows/columns in one call. The worksheet tracks the groups and computes outline levels automatically. |
| **Low-level** — `set_outline_level` / `set_collapsed` on each `zcl_excel_row` or `zcl_excel_column` | Fine-grained control when you need to assign outline levels row-by-row or when you are reconstructing outline state from a data model. |

Both approaches produce valid OOXML output. The high-level approach is simpler for the common case.

---

## High-Level Approach: `add_row_outline` / `add_column_outline`

### Row grouping

```abap
" Group rows 3 to 7 (level 1 outline — expandable block)
lo_worksheet->add_row_outline(
  ip_row_from = 3
  ip_row_to   = 7
).

" Nest a sub-group inside it (rows 4 to 5 become level 2)
lo_worksheet->add_row_outline(
  ip_row_from = 4
  ip_row_to   = 5
).

" Collapse the outer group so rows 3-7 are hidden on open
lo_worksheet->add_row_outline(
  ip_row_from  = 3
  ip_row_to    = 7
  ip_collapsed = abap_true
).
```

### Column grouping

```abap
" Group columns B to D (outline level 1)
lo_worksheet->add_column_outline(
  ip_column_from = 'B'
  ip_column_to   = 'D'
).

" Collapse it on open
lo_worksheet->add_column_outline(
  ip_column_from = 'B'
  ip_column_to   = 'D'
  ip_collapsed   = abap_true
).
```

### Reading back outline groups

```abap
DATA lt_row_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_row.
lt_row_outlines = lo_worksheet->get_row_outlines( ).
" Each entry: row_from, row_to, collapsed (abap_bool)

DATA lt_col_outlines TYPE zcl_excel_worksheet=>mty_ts_outlines_column.
lt_col_outlines = lo_worksheet->get_column_outlines( ).
" Each entry: column_from (alpha), column_to (alpha), collapsed (abap_bool)
```

---

## Low-Level Approach: Row and Column Objects

When you need per-row or per-column control, work directly with `zcl_excel_row` and `zcl_excel_column` objects retrieved from the worksheet.

### Row objects

```abap
DATA lo_row TYPE REF TO zcl_excel_row.

" Get (creates if absent) the row object for row 5
lo_row = lo_worksheet->get_row( 5 ).

" Set this row as outline level 1 (detail row inside a group)
lo_row->set_outline_level( 1 ).

" Hide the row (collapsed group member)
lo_row->set_visible( abap_false ).

" Set row height in points
lo_row->set_row_height( 20 ).
```

### Column objects

```abap
DATA lo_column TYPE REF TO zcl_excel_column.

" Get (creates if absent) the column object for column C
lo_column = lo_worksheet->get_column( 'C' ).

" Set as outline level 1 detail column
lo_column->set_outline_level( 1 ).

" Hide it (collapsed)
lo_column->set_visible( abap_false ).

" Mark the group header column as collapsed
lo_column->set_collapsed( abap_true ).
```

### Outline level rules

- Valid outline levels: **0** (no grouping) to **7** (deepest nesting).
- `set_outline_level` on `zcl_excel_row` raises `zcx_excel` if the value is outside 0–7.
- `set_outline_level` on `zcl_excel_column` does not validate the range — ensure you stay within 0–7 to produce valid OOXML.
- When `set_collapsed( abap_true )` is set on a row/column object, Excel renders the **expand/collapse button** on the summary row/column.

---

## Worked Example: Collapsible Detail Rows

The following example builds a report with summary rows followed by detail rows. The detail rows are grouped and collapsed by default so the sheet opens in a compact view.

```abap
DATA lo_row TYPE REF TO zcl_excel_row.

" --- Summary rows (level 0 — always visible) ---
lo_worksheet->set_cell( ip_row = 1  ip_column = 'A'  ip_value = 'Region: EMEA' ).
lo_worksheet->set_cell( ip_row = 5  ip_column = 'A'  ip_value = 'Region: APAC' ).

" --- Detail rows for EMEA (rows 2-4) ---
lo_worksheet->set_cell( ip_row = 2  ip_column = 'A'  ip_value = 'Germany' ).
lo_worksheet->set_cell( ip_row = 3  ip_column = 'A'  ip_value = 'France' ).
lo_worksheet->set_cell( ip_row = 4  ip_column = 'A'  ip_value = 'UK' ).

" Group rows 2-4 as outline level 1 and collapse them
lo_worksheet->add_row_outline(
  ip_row_from  = 2
  ip_row_to    = 4
  ip_collapsed = abap_true
).

" Detail rows for APAC (rows 6-7)
lo_worksheet->set_cell( ip_row = 6  ip_column = 'A'  ip_value = 'Japan' ).
lo_worksheet->set_cell( ip_row = 7  ip_column = 'A'  ip_value = 'Australia' ).

" Group rows 6-7 but leave them expanded
lo_worksheet->add_row_outline(
  ip_row_from  = 6
  ip_row_to    = 7
  ip_collapsed = abap_false
).
```

When this file opens in Excel, the EMEA detail rows (2–4) are hidden and there is an expand button `[+]` next to row 1. The APAC detail rows (6–7) are visible and a collapse button `[-]` appears next to row 5.

---

## Summary Direction: `summaryBelow` and `summaryRight`

By default Excel places the summary row **below** the detail rows and the summary column **to the right** of the detail columns. You can reverse this through the worksheet sheet properties:

```abap
" Place summary row ABOVE the detail rows (SAP-style totals at top)
lo_worksheet->zif_excel_sheet_properties~summarybelow
  = zif_excel_sheet_properties=>c_below_off.   " abap_false

" Place summary column to the LEFT of detail columns
lo_worksheet->zif_excel_sheet_properties~summaryright
  = zif_excel_sheet_properties=>c_right_off.   " abap_false
```

::: tip SAP-style totals
SAP ALV grid places totals rows **above** detail lines. If you are exporting ALV data with grouping, set `summarybelow = c_below_off` so the collapse/expand buttons appear on the correct side.
:::

The `summarybelow` flag also affects which row receives the `collapsed` flag in the OOXML when using the high-level `add_row_outline` API — the `get_collapsed` method on `zcl_excel_row` consults this flag when determining whether a row is hidden inside a collapsed group.

---

## Column Width and Style on Grouped Columns

When you retrieve a column object you can also set its width and a column-wide style in the same chain:

```abap
lo_column = lo_worksheet->get_column( 'B' ).
lo_column->set_outline_level( 1 ).
lo_column->set_width( 18 ).             " width in character units
lo_column->set_auto_size( abap_false ).

" Apply a column-wide number format / style
DATA(lv_style_guid) = lo_excel->add_new_style( ).
" ... configure the style object ...
lo_column->set_column_style_by_guid( lv_style_guid ).
```

`set_column_style_by_guid` falls back to the worksheet default style if the supplied GUID is initial or invalid, so it is safe to call even when no style has been explicitly created.

---

## Nesting Outlines (Multi-Level)

Outlines can be nested up to 7 levels deep. Each call to `add_row_outline` adds one nesting level on top of any existing groups that overlap the same range.

```abap
" Outer group: rows 2-10  (level 1)
lo_worksheet->add_row_outline( ip_row_from = 2  ip_row_to = 10 ).

" Inner group: rows 3-5  (level 2 — overlaps with the outer group)
lo_worksheet->add_row_outline( ip_row_from = 3  ip_row_to = 5 ).

" Deepest group: row 4 only  (level 3)
lo_worksheet->add_row_outline( ip_row_from = 4  ip_row_to = 4 ).
```

Excel displays three levels of outline controls `[1] [2] [3]` in the top-left corner of the sheet.

---

## Cloud Compatibility

`zcl_excel_row`, `zcl_excel_column`, and the worksheet outline methods are all in the main `src/` package and use no restricted ABAP statements. They are fully cloud-compatible and work on SAP BTP ABAP Environment without modification.

---

## Limitations

- Excel limits outlines to **8 levels** (0–7). abap2xlsx enforces this at the row level (`zcx_excel` raised on violation); validate manually for columns.
- Reading outline groups back from an `.xlsx` file produced by Excel or a third-party tool is not currently supported — `zcl_excel_reader_2007` does not parse `<row outlineLevel>` attributes into `add_row_outline` entries.
- `get_row_outlines` / `get_column_outlines` only return groups registered through the high-level `add_row_outline` / `add_column_outline` API in the current session. Outline levels set via the low-level `set_outline_level` API are **not** reflected in those collections.

---

## Next Steps

- [Worksheets](./worksheets.md) — freeze panes, sheet protection, print settings
- [Page Breaks](./worksheets.md#page-breaks) — row and column page break insertion
- [Autofilter](./autofilter.md) — combining autofilter with grouped rows
- [Formatting](./formatting.md) — applying styles to rows and columns
