# AutoFilter

`zcl_excel_autofilter` adds an Excel AutoFilter to a worksheet — the drop-down filter arrows
in a header row that let users show or hide rows interactively. The same class is also used
by `zcl_excel_writer_csv` to honour active filters when exporting to CSV.

## Adding an AutoFilter to a Worksheet

Each worksheet owns exactly one autofilter object. Retrieve it with `get_autofilter()`,
configure it, and the writer serialises it automatically:

```abap
DATA: lo_autofilter TYPE REF TO zcl_excel_autofilter.

" Retrieve (or lazily create) the worksheet's autofilter object
lo_autofilter = lo_worksheet->get_autofilter( ).
```

### Setting the Filter Area

If you skip this step the engine defaults to the worksheet's full data range. For precise
control, set the area explicitly:

```abap
" Define the filter area — row_start is the header row, row_end is the last data row
lo_autofilter->set_filter_area(
  VALUE zexcel_s_autofilter_area(
    row_start = 1
    col_start = 1
    row_end   = 100
    col_end   = 5
  )
).
```

The `validate_area` method (called internally) auto-corrects out-of-range values against the
worksheet's actual data extent, so you do not need to know the exact row count upfront.

## Filtering by Single Values

Use `set_value` to show only rows where a column's cell equals a specific value.
Multiple calls on the **same column** accumulate into a union (logical OR).
Multiple calls on **different columns** act as AND conditions:

```abap
" Show only rows where column 1 (A) = 'North' OR 'South'
lo_autofilter->set_value( i_column = 1  i_value = 'North' ).
lo_autofilter->set_value( i_column = 1  i_value = 'South' ).

" AND column 2 (B) = 'Active'
lo_autofilter->set_value( i_column = 2  i_value = 'Active' ).
```

For bulk loading from an internal table use `set_values`:

```abap
DATA lt_values TYPE zexcel_t_autofilter_values.
APPEND VALUE #( column = 1  value = 'North'  ) TO lt_values.
APPEND VALUE #( column = 1  value = 'South'  ) TO lt_values.
APPEND VALUE #( column = 2  value = 'Active' ) TO lt_values.

lo_autofilter->set_values( lt_values ).
```

## Filtering by Text Pattern

`set_text_filter` applies a wildcard or exact-match pattern to a column.
Use `*` and `+` as wildcards (same semantics as ABAP `CP` operator):

```abap
" Constant for filter rule type
" zcl_excel_autofilter=>mc_filter_rule_text_pattern

" Show rows where column 3 starts with 'SAP'
lo_autofilter->set_text_filter(
  i_column       = 3
  iv_textfilter1 = 'SAP*'   " * = any characters
).

" Exact match (no wildcards) — equivalent to set_value
lo_autofilter->set_text_filter(
  i_column       = 3
  iv_textfilter1 = 'SAP SE'
).
```

> **Note:** `set_text_filter` replaces any previous filter on that column. Only one text
> pattern per column is supported by the current implementation.

## Filter Rule Constants

| Constant | Value | Used by |
|---|---|---|
| `mc_filter_rule_single_values` | `'single_values'` | `set_value` / `set_values` |
| `mc_filter_rule_text_pattern` | `'text_pattern'` | `set_text_filter` |
| `mc_logical_operator_and` | `'and'` | Internal (future use) |
| `mc_logical_operator_or` | `'or'` | Internal (future use) |
| `mc_logical_operator_none` | `space` | Default |

## Reading Filter State (for CSV Export)

`is_row_hidden( iv_row )` returns `abap_true` if a row would be hidden by the active
filter rules. `zcl_excel_writer_csv` calls this automatically when its
`mv_skip_hidden_rows` flag is set. You can also call it directly:

```abap
DATA(lv_row) = 2.
WHILE lv_row <= lo_worksheet->get_highest_row( ).
  IF lo_autofilter->is_row_hidden( lv_row ) = abap_false.
    " Process this row
  ENDIF.
  ADD 1 TO lv_row.
ENDWHILE.
```

The first row of the filter area (the header row) is **never** considered hidden, regardless
of filter values.

## Getting the Filter Range as a String

Two helper methods return the filter area as a string — useful for debugging or for passing
to other methods:

```abap
" Returns e.g. 'A1:E100'
DATA(lv_range) = lo_autofilter->get_filter_range( ).

" Returns a fully-qualified reference, e.g. 'Sales!$A$1:$E$100'
DATA(lv_ref) = lo_autofilter->get_filter_reference( ).
```

## Full Example

```abap
" Create workbook with data
DATA(lo_excel)     = NEW zcl_excel( ).
DATA(lo_worksheet) = lo_excel->add_new_worksheet( ).
lo_worksheet->set_title( 'Sales Data' ).

" Write header row
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Region' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Product' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Status' ).

" Write data rows
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'North' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Laptop' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = 'Active' ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'South' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Monitor' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Active' ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = 'North' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Keyboard' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 4 ip_value = 'Inactive' ).

" Add AutoFilter covering all 3 columns and 4 rows (header + 3 data rows)
DATA(lo_af) = lo_worksheet->get_autofilter( ).
lo_af->set_filter_area(
  VALUE zexcel_s_autofilter_area(
    row_start = 1  col_start = 1
    row_end   = 4  col_end   = 3
  )
).

" Filter: region must be 'North' AND status must start with 'Act'
lo_af->set_value( i_column = 1  i_value = 'North' ).
lo_af->set_text_filter( i_column = 3  iv_textfilter1 = 'Act*' ).

" Export — the xlsx file will have filter arrows in row 1;
" rows 3 and 4 are hidden by the active filter.
DATA(lo_writer) = NEW zcl_excel_writer_2007( ).
DATA(lv_file)   = lo_writer->write_file( lo_excel ).
```

## AutoFilter and CSV Export

When exporting to CSV via `zcl_excel_writer_csv`, set `mv_skip_hidden_rows = abap_true`
to omit rows that the active AutoFilter would hide:

```abap
DATA(lo_csv) = NEW zcl_excel_writer_csv( ).
lo_csv->mv_skip_hidden_rows    = abap_true.
lo_csv->mv_skip_hidden_columns = abap_true.  " Also skip hidden columns
DATA(lv_csv) = lo_csv->write_file( lo_excel ).
```

See **[CSV Export](/guide/csv-export)** for full options.

## Next Steps

- **[CSV Export](/guide/csv-export)** — skip hidden rows/columns when exporting
- **[Worksheets](/guide/worksheets)** — freeze panes, data validation, page breaks
- **[Data Conversion](/guide/data-conversion)** — `bind_table` with built-in table styles
- **[Changelog](/guide/changelog)** — history of autofilter changes
