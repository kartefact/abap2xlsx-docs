# Worksheets

A worksheet (`zcl_excel_worksheet`) is the primary canvas for all cell data, styling, drawing, and sheet-level configuration in abap2xlsx. Every `zcl_excel` workbook holds one or more worksheets.

## Creating and accessing worksheets

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

CREATE OBJECT lo_excel.

" First worksheet is created automatically
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_title( 'Sales Data' ).

" Add further worksheets
lo_worksheet = lo_excel->add_new_worksheet( ).
lo_worksheet->set_title( 'Summary' ).
```

## Writing cells

`set_cell` is the primary method for writing a single cell. It accepts the address either as an alpha reference (`ip_columnrow`) or as separate column + row integers.

```abap
" Alpha reference
lo_worksheet->set_cell( ip_columnrow = 'B3' ip_value = 'Hello' ).

" Column + row integers
lo_worksheet->set_cell( ip_column = 2 ip_row = 3 ip_value = 'Hello' ).

" With a style GUID
lo_worksheet->set_cell(
  ip_column   = 2
  ip_row      = 3
  ip_value    = lv_amount
  ip_style    = lv_style_guid ).

" With a formula
lo_worksheet->set_cell(
  ip_column  = 5
  ip_row     = 10
  ip_formula = 'SUM(E2:E9)' ).
```

### ABAP type mapping

`set_cell` inspects the runtime type of `ip_value` and maps it to the correct Excel data type automatically.

| ABAP type | Excel type written | Notes |
|---|---|---|
| `I`, `F`, `P`, `DECFLOAT16/34` | Number | Decimal separator adjusted |
| `D` (date) | Number (date serial) | Uses worksheet date format |
| `T` (time) | Number (time fraction) | Stored as fractional day |
| `UTCLONG` | Number (datetime serial) | S/4HANA only — see [Data Conversion](./data-conversion.md) |
| `C`, `N`, `STRING` | String or Number | Leading-zero strings stay as text |
| `X` (XSTRING) | Not directly supported | Convert to base64 or use a drawing |

## Reading cells

```abap
DATA: lv_value   TYPE zexcel_cell_value,
      lv_formula TYPE zexcel_cell_formula,
      lv_rc      TYPE sysubrc.

lo_worksheet->get_cell(
  ip_column  = 3
  ip_row     = 5
  IMPORTING
    ep_value   = lv_value
    ep_formula = lv_formula
    ep_rc      = lv_rc ).  " sy-subrc: 0 = found, 4 = empty
```

## Cell ranges and bulk operations

`set_area` writes the same value, formula, style, or hyperlink to every cell in a rectangular range.

```abap
" Fill a range by alpha reference
lo_worksheet->set_area(
  ip_range   = 'A1:D1'
  ip_style   = lv_header_style ).

" Fill a range by column/row bounds
lo_worksheet->set_area(
  ip_column_start = 1
  ip_column_end   = 4
  ip_row          = 1
  ip_row_to       = 1
  ip_value        = 'Header'
  ip_style        = lv_style ).
```

Related range methods:

| Method | Purpose |
|---|---|
| `set_area_style` | Apply a style to all cells in a range |
| `set_area_formula` | Write one formula to all cells in a range |
| `set_area_hyperlink` | Attach a hyperlink to all cells in a range |
| `change_area_style` | Modify part of the style on an existing range via `zif_excel_style_changer` |
| `set_merge` | Merge a range of cells (optionally with a value and formula) |
| `set_merge_style` | Apply a style to a merged range |
| `delete_merge` | Remove a cell merge |
| `get_merge` | Return a string table of all current merge ranges |
| `is_cell_merged` | Check if a specific cell is part of a merge |

## Column and row sizing

```abap
" Fixed pixel width
lo_worksheet->set_column_width(
  ip_column    = 3
  ip_width_fix = 20 ).

" Auto-size (measure cell content)
lo_worksheet->set_column_width(
  ip_column         = 3
  ip_width_autosize = abap_true ).

" Recalculate all column widths based on content
lo_worksheet->calculate_column_widths( ).

" Fixed row height
lo_worksheet->set_row_height(
  ip_row        = 1
  ip_height_fix = 25 ).
```

## Freeze panes

`freeze_panes` locks the specified number of columns and/or rows so they remain visible while scrolling.

```abap
" Freeze top row
lo_worksheet->freeze_panes( ip_num_rows = 1 ).

" Freeze first two columns
lo_worksheet->freeze_panes( ip_num_columns = 2 ).

" Freeze both rows and columns
lo_worksheet->freeze_panes( ip_num_columns = 1 ip_num_rows = 3 ).
```

### Pane scroll position

After freezing panes, you can control which cell appears in the top-left corner of the scrollable area, and which cell is the initial scroll position of the full sheet view.

```abap
" After freeze_panes( ip_num_rows = 3 ), show column D row 4 in the scrollable pane
lo_worksheet->set_pane_top_left_cell( iv_columnrow = 'D4' ).

" Set the initial top-left visible cell when the sheet is opened (no freeze required)
lo_worksheet->set_sheetview_top_left_cell( iv_columnrow = 'F10' ).
```

- `set_pane_top_left_cell` — controls the top-left cell of the **frozen pane's scrollable region**. Useful when you want the sheet to open pre-scrolled past a large header block.
- `set_sheetview_top_left_cell` — sets the top-left cell of the **entire sheet view** when first opened. Does not require a freeze pane to be active.

Both methods accept any valid alpha cell reference (`'A1'`, `'D4'`, `'Z100'`, etc.).

## Tab colour

```abap
DATA: ls_color TYPE zexcel_s_tabcolor.
ls_color-rgb = 'FF4A90E2'.  " ARGB: opaque blue
lo_worksheet->set_tabcolor( ls_color ).
```

## Grid and header visibility

```abap
lo_worksheet->set_show_gridlines( abap_false ).  " Hide gridlines
lo_worksheet->set_show_rowcolheaders( abap_false ).  " Hide row/column headers
lo_worksheet->set_print_gridlines( abap_true ).  " Print gridlines
```

## Dimension range

`get_dimension_range` returns the used cell range of the worksheet as an Excel reference string (e.g. `'A1:G42'`):

```abap
DATA(lv_range) = lo_worksheet->get_dimension_range( ).
```

This is recalculated automatically as cells are written.

## Highest used row / column

```abap
DATA(lv_max_row)    = lo_worksheet->get_highest_row( ).
DATA(lv_max_column) = lo_worksheet->get_highest_column( ).
```

## Named ranges

```abap
DATA: lo_range TYPE REF TO zcl_excel_range.
lo_range = lo_worksheet->add_new_range( ).
lo_range->set_name( 'MyRange' ).
lo_range->set_value( 'Sheet1!$A$1:$D$10' ).
```

See [Named Ranges](./named-ranges.md) for the full guide.

## Row grouping and outlines

Row grouping (also called outlining) adds the Excel expand/collapse controls to the left of a sheet. This is implemented via `set_row_outline`, `delete_row_outline`, and `get_row_outlines` on the worksheet.

```abap
" Group rows 5 through 12, initially collapsed
lo_worksheet->set_row_outline(
  iv_row_from  = 5
  iv_row_to    = 12
  iv_collapsed = abap_true ).

" Group rows 5 through 12, expanded
lo_worksheet->set_row_outline(
  iv_row_from  = 5
  iv_row_to    = 12
  iv_collapsed = abap_false ).

" Remove a row outline group
lo_worksheet->delete_row_outline(
  iv_row_from = 5
  iv_row_to   = 12 ).

" Read all current outline groups
DATA(lt_outlines) = lo_worksheet->get_row_outlines( ).
```

The internal table returned by `get_row_outlines` uses type `mty_ts_outlines_row`, a sorted table with unique key on `row_from`/`row_to`. Each entry has:

| Field | Type | Meaning |
|---|---|---|
| `row_from` | `i` | First row in the group |
| `row_to` | `i` | Last row in the group |
| `collapsed` | `abap_bool` | `abap_true` = group is collapsed on open |

### Nested outline groups

Excel supports up to 8 levels of nesting. Create nested groups by adding multiple overlapping or contained ranges:

```abap
" Outer group: rows 2–20
lo_worksheet->set_row_outline( iv_row_from = 2  iv_row_to = 20 iv_collapsed = abap_false ).
" Inner group: rows 5–10
lo_worksheet->set_row_outline( iv_row_from = 5  iv_row_to = 10 iv_collapsed = abap_true ).
" Another inner group: rows 14–18
lo_worksheet->set_row_outline( iv_row_from = 14 iv_row_to = 18 iv_collapsed = abap_true ).
```

> **Column outlines** are not yet supported by the writer. Use the `bind_alv` path with an ALV layout that has column grouping if you need column-level outlining.

## Page breaks

```abap
DATA(lo_pb) = lo_worksheet->get_pagebreaks( ).

" Insert a row page break after row 40
lo_pb->add_pagebreak( ip_row = 40 ).

" Insert a column page break after column 8 (column H)
lo_pb->add_pagebreak( ip_column = 8 ).

" Programmatic page break every 50 rows
DO.
  DATA(lv_break_row) = sy-index * 50.
  IF lv_break_row > lv_last_data_row. EXIT. ENDIF.
  lo_pb->add_pagebreak( ip_row = lv_break_row ).
ENDDO.
```

See [zcl_excel_worksheet_pagebreaks](./worksheets.md) constants:
- `zcl_excel_worksheet=>c_break_row` (`1`) — row break
- `zcl_excel_worksheet=>c_break_column` (`2`) — column break
- `zcl_excel_worksheet=>c_break_none` (`0`) — no break

## Suppressing cell validation warnings

Use `set_ignored_errors` to suppress Excel's green-triangle warnings (numbers as text, formula deviations, etc.) on specific cell ranges. See the dedicated [Ignored Errors](./ignored-errors.md) guide for all ten available flags.

```abap
DATA: lt_ie TYPE zcl_excel_worksheet=>mty_th_ignored_errors,
      ls_ie TYPE zcl_excel_worksheet=>mty_s_ignored_errors.

ls_ie-cell_coords          = 'A2:A500'.
ls_ie-number_stored_as_text = abap_true.
INSERT ls_ie INTO TABLE lt_ie.
lo_worksheet->set_ignored_errors( lt_ie ).
```

## Active cell

```abap
DATA(lv_active) = lo_worksheet->get_active_cell( ).
```

Returns the currently selected cell reference (e.g. `'B3'`).

## Print settings

Print configuration is handled through the `zif_excel_sheet_printsettings` interface, which `zcl_excel_worksheet` implements:

```abap
lo_worksheet->zif_excel_sheet_printsettings~set_paper_size(
  zcl_excel_sheet_setup=>c_paper_a4 ).
lo_worksheet->zif_excel_sheet_printsettings~set_orientation(
  zcl_excel_sheet_setup=>c_orientation_landscape ).
lo_worksheet->zif_excel_sheet_printsettings~set_scale( 85 ).  " 85%
lo_worksheet->zif_excel_sheet_printsettings~set_fittopage( abap_true ).
```

## Sheet properties

Available through the `zif_excel_sheet_properties` interface:

```abap
" Set right-to-left display for Arabic / Hebrew sheets
" (no public setter on zcl_excel_worksheet — configure via sheet_setup)
lo_worksheet->sheet_setup->set_tab_color(
  ip_color_rgb = 'FF00AA00' ).
```

## Sheet protection

See [Workbook Security](./workbook-security.md).

## Converting worksheet data back to ABAP

To read cell data from an already-loaded worksheet back into an ABAP internal table, use `convert_to_table`. See [Reading Worksheet Data](./convert-to-table.md) for the full guide, including the difference between typed (`et_data`) and lossless string (`er_data`) output.

```abap
lo_worksheet->convert_to_table(
  EXPORTING
    it_field_catalog = lt_catalog
    iv_begin_row     = 2
  IMPORTING
    et_data          = lt_result ).
```

## Iterators

All sub-collections on a worksheet are accessible through iterator objects:

| Method | Returns |
|---|---|
| `get_columns_iterator` | Iterator over `zcl_excel_column` objects |
| `get_rows_iterator` | Iterator over `zcl_excel_row` objects |
| `get_comments_iterator` | Iterator over `zcl_excel_comment` objects |
| `get_drawings_iterator(ip_type)` | Iterator over charts or images |
| `get_hyperlinks_iterator` | Iterator over `zcl_excel_hyperlink` objects |
| `get_ranges_iterator` | Iterator over `zcl_excel_range` objects |
| `get_tables_iterator` | Iterator over `zcl_excel_table` objects |
| `get_style_cond_iterator` | Iterator over conditional style objects |
| `get_data_validations_iterator` | Iterator over data validation objects |

All iterators follow the same pattern:

```abap
DATA: lo_iter TYPE REF TO zcl_excel_collection_iterator,
      lo_obj  TYPE REF TO object.

lo_iter = lo_worksheet->get_comments_iterator( ).
WHILE lo_iter->has_next( ) = abap_true.
  lo_obj = lo_iter->get_next( ).
  " Cast and use lo_obj
ENDWHILE.
```

## See also

- [Formatting](./formatting.md) — styles, fonts, fills, borders
- [Cell Comments](./cell-comments.md) — adding comments to cells
- [Ignored Errors](./ignored-errors.md) — suppressing green-triangle warnings
- [Convert to Table](./convert-to-table.md) — reading worksheet data back into ABAP
- [Row/Column Grouping](./row-column-grouping.md) — column outlines and grouping
- [Autofilter](./autofilter.md) — dropdown column filters
- [Named Ranges](./named-ranges.md) — workbook-level named ranges
- [Data Validation](./data-validation.md) — constraining cell input
- [Template Filling](./template-filling.md) — named-range token substitution
