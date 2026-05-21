# Excel Tables

`zcl_excel_table` creates an Excel **structured table** — the rectangular range with banded
rows, column header filters, and optional totals row that Excel calls a "Table" (Insert →
Table, keyboard `Ctrl+T`). Structured tables are distinct from plain data ranges: they carry
a name, a style, and support structured-reference formulas like `=SUM(Sales[Amount])`.

## Creating a Table

Tables are registered at the **worksheet** level via `add_new_table()`:

```abap
DATA: lo_table TYPE REF TO zcl_excel_table.

" add_new_table creates the table object and registers it on the worksheet
lo_table = lo_worksheet->add_new_table( ).
```

## Configuring the Table

### Name and Range

```abap
" Table name — used in structured references like =SUM(SalesData[Amount])
" Must be unique within the workbook and match the Excel name rules
" (letters, digits, underscore — no spaces)
lo_table->set_name( 'SalesData' ).

" Range — must include the header row
" Tip: use zcl_excel_common=>convert_column2alpha() for column letters
lo_table->set_ref( 'A1:E101' ).  " Header in row 1, data in rows 2-101
```

### Table Style

```abap
" Built-in table style names follow the pattern TableStyleLight/Medium/Dark + number
" Light: TableStyleLight1 .. TableStyleLight21
" Medium: TableStyleMedium1 .. TableStyleMedium28
" Dark: TableStyleDark1 .. TableStyleDark11
lo_table->set_style( zcl_excel_table=>c_style_medium2 ).  " Popular blue style

" Fine-grained visibility flags
lo_table->set_show_first_column( abap_false ).  " First column special formatting
lo_table->set_show_last_column( abap_false ).   " Last column special formatting
lo_table->set_show_row_stripes( abap_true ).    " Alternate row shading (banded rows)
lo_table->set_show_column_stripes( abap_false ). " Alternate column shading
```

**Style constants on `zcl_excel_table`:**

| Constant | Excel style name |
|---|---|
| `c_style_light1` … `c_style_light21` | `TableStyleLight1` … `TableStyleLight21` |
| `c_style_medium1` … `c_style_medium28` | `TableStyleMedium1` … `TableStyleMedium28` |
| `c_style_dark1` … `c_style_dark11` | `TableStyleDark1` … `TableStyleDark11` |
| `c_style_none` | `TableStyleNone` (no style) |

### Header Row

```abap
" The header row is shown by default — the top row of the table range
" is treated as column headers. Hiding it is rare and not recommended.
lo_table->set_header_row_count( 1 ).  " 1 = default: one header row visible
```

### AutoFilter

```abap
" AutoFilter (the drop-down arrows in the header row) is ON by default
" for all structured tables. To disable:
lo_table->set_auto_filter( abap_false ).
```

### Totals Row

```abap
" Enable a totals row at the bottom of the table
lo_table->set_totals_row_shown( abap_true ).

" Configure the aggregation function for each column
" Column index is 1-based within the table, not the sheet column number
DATA: lo_column TYPE REF TO zcl_excel_table_column.

" Get or create a column descriptor by 1-based table-column index
lo_column = lo_table->get_column( 1 ).  " First table column
lo_column->set_name( 'Region' ).        " Column header label
lo_column->set_totals_row_label( 'Total' ).  " Label shown in the totals row cell

" Column 5: Sum the Amount column
lo_column = lo_table->get_column( 5 ).
lo_column->set_name( 'Amount' ).
lo_column->set_totals_row_function( zcl_excel_table_column=>c_totals_row_function_sum ).
```

**`totalsRowFunction` constants on `zcl_excel_table_column`:**

| Constant | Excel formula | Notes |
|---|---|---|
| `c_totals_row_function_sum` | `SUBTOTAL(109,...)` | Sum of visible rows |
| `c_totals_row_function_count` | `SUBTOTAL(102,...)` | Count of non-empty visible rows |
| `c_totals_row_function_countNums` | `SUBTOTAL(102,...)` | Count of numeric values |
| `c_totals_row_function_average` | `SUBTOTAL(101,...)` | Average of visible rows |
| `c_totals_row_function_max` | `SUBTOTAL(104,...)` | Maximum value |
| `c_totals_row_function_min` | `SUBTOTAL(105,...)` | Minimum value |
| `c_totals_row_function_stdDev` | `SUBTOTAL(107,...)` | Standard deviation |
| `c_totals_row_function_var` | `SUBTOTAL(110,...)` | Variance |
| `c_totals_row_function_none` | *(no formula)* | Plain label cell |
| `c_totals_row_function_custom` | Any formula string | Use `set_totals_row_formula` |

> **Reader support:** `zcl_excel_reader_2007` reads `totalsRowFunction` from existing
> `.xlsx` files as of the Feb 2026 update (PR [#1296](https://github.com/abap2xlsx/abap2xlsx/pull/1296)).

### Custom Totals Formula

```abap
" For custom aggregations, set function = custom and provide the formula
lo_column->set_totals_row_function(
  zcl_excel_table_column=>c_totals_row_function_custom
).
lo_column->set_totals_row_formula(
  'SUM(SalesData[Amount])/COUNT(SalesData[Region])'
).
```

## Full Example: Sales Table with Totals Row

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_table     TYPE REF TO zcl_excel_table,
      lo_col       TYPE REF TO zcl_excel_table_column.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_title( 'Sales' ).

" Write column headers
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Region'  ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Product' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Qty'     ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Price'   ).
lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Amount'  ).

" Write data rows
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'North'  ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Laptop' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = 10       ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = 1200     ).
lo_worksheet->set_cell( ip_column = 'E' ip_row = 2 ip_value = 12000    ).

lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'South'   ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Monitor' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 5         ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = 400       ).
lo_worksheet->set_cell( ip_column = 'E' ip_row = 3 ip_value = 2000      ).

" Create and configure the structured table
lo_table = lo_worksheet->add_new_table( ).
lo_table->set_name( 'SalesData' ).
lo_table->set_ref( 'A1:E3' ).
lo_table->set_style( zcl_excel_table=>c_style_medium2 ).
lo_table->set_show_row_stripes( abap_true ).
lo_table->set_totals_row_shown( abap_true ).

" Column headers and totals functions
lo_col = lo_table->get_column( 1 ). lo_col->set_name( 'Region'  ).
  lo_col->set_totals_row_label( 'Total' ).
lo_col = lo_table->get_column( 2 ). lo_col->set_name( 'Product' ).
lo_col = lo_table->get_column( 3 ). lo_col->set_name( 'Qty'     ).
  lo_col->set_totals_row_function( zcl_excel_table_column=>c_totals_row_function_sum ).
lo_col = lo_table->get_column( 4 ). lo_col->set_name( 'Price'   ).
lo_col = lo_table->get_column( 5 ). lo_col->set_name( 'Amount'  ).
  lo_col->set_totals_row_function( zcl_excel_table_column=>c_totals_row_function_sum ).

" Serialise
DATA(lo_writer) = NEW zcl_excel_writer_2007( ).
DATA(lv_file)   = lo_writer->write_file( lo_excel ).
```

## Using `bind_table` with a Structured Table

`bind_table` can populate the data and simultaneously register a structured table:

```abap
" bind_table with ip_table_style creates a structured table automatically
lo_worksheet->bind_table(
  ip_table        = lt_sales_data
  ip_table_style  = zcl_excel_table=>c_style_medium2
  ip_header_row   = 1
  ip_start_column = 1
  ip_start_row    = 1
).
```

See **[Data Conversion](/guide/data-conversion)** for the full `bind_table` parameter reference.

## Structured Reference Formulas

Once a table is named, standard Excel structured reference syntax works in formulas:

```abap
" Sum all amounts in the SalesData table
lo_worksheet->set_cell_formula(
  ip_column  = 'G'
  ip_row     = 1
  ip_formula = 'SUM(SalesData[Amount])'
).

" Count rows where Region = 'North'
lo_worksheet->set_cell_formula(
  ip_column  = 'G'
  ip_row     = 2
  ip_formula = 'COUNTIF(SalesData[Region],"North")'
).
```

## Next Steps

- **[AutoFilter](/guide/autofilter)** — add column drop-down filters to plain ranges
- **[Data Conversion](/guide/data-conversion)** — `bind_table` for quick table population
- **[Reading Excel](/guide/reading-excel)** — read `totalsRowFunction` from existing files
- **[Formatting](/guide/formatting)** — apply cell styles to table data
