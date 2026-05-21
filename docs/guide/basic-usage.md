# Basic Usage

Comprehensive guide to creating your first Excel files with abap2xlsx.

## Core Concepts

### The Excel Object Model

abap2xlsx follows Excel's object hierarchy:

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer    TYPE REF TO zif_excel_writer.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).
CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
DATA(lv_file) = lo_writer->write_file( lo_excel ).
```

### Understanding Cell References

```abap
lo_worksheet->set_cell( ip_column = 'A'  ip_row = 1 ip_value = 'Cell A1' ).
lo_worksheet->set_cell( ip_column = 'B'  ip_row = 1 ip_value = 'Cell B1' ).
lo_worksheet->set_cell( ip_column = 'AA' ip_row = 1 ip_value = 'Cell AA1' ).
DATA(lv_value) = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).
```

## Creating Your First Workbook

```abap
REPORT zcreate_first_workbook.

DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer    TYPE REF TO zif_excel_writer.

START-OF-SELECTION.
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->get_active_worksheet( ).
  lo_worksheet->set_title( 'My Data' ).

  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Product' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Quantity' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Price' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Laptop' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 10 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = '999.99' ).

  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_excel_file) = lo_writer->write_file( lo_excel ).
  MESSAGE 'Excel file created successfully' TYPE 'S'.
```

## Working with Data Types

### ABAP to Excel Data Type Mapping

```abap
DATA: lv_string  TYPE string    VALUE 'Text Value',
      lv_integer TYPE i         VALUE 42,
      lv_decimal TYPE p DECIMALS 2 VALUE '123.45',
      lv_date    TYPE d         VALUE '20231225',
      lv_time    TYPE t         VALUE '143000',
      lv_boolean TYPE abap_bool VALUE abap_true.

lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = lv_string ).   " Text
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = lv_integer ).  " Number
lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = lv_decimal ).  " Decimal
lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = lv_date ).     " Date
lo_worksheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = lv_time ).     " Time
lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = lv_boolean ).  " Boolean
```

### set_cell XSTRING Support

> **Added in May 2025 (PR [#1306](https://github.com/abap2xlsx/abap2xlsx/pull/1306)):** `set_cell()` now accepts `XSTRING`-typed values directly.

```abap
DATA lv_xstr TYPE xstring.
" ... populate lv_xstr from a stream or function module ...

lo_worksheet->set_cell(
  ip_column = 'A'
  ip_row    = 1
  ip_value  = lv_xstr   " XSTRING now accepted directly
).
```

The value is stored as-is in the cell's internal representation. When reading the file back, retrieve the cell value as a string and apply your own hex-to-binary conversion if needed.

## Working with Internal Tables

### Manual Loop

```abap
TYPES: BEGIN OF ty_sales,
         region    TYPE string,
         product   TYPE string,
         quantity  TYPE i,
         revenue   TYPE p DECIMALS 2,
         sale_date TYPE d,
       END OF ty_sales.
DATA lt_sales TYPE TABLE OF ty_sales.

" Headers
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Region' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Product' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Quantity' ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Revenue' ).
lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Sale Date' ).

DATA(lv_row) = 2.
LOOP AT lt_sales INTO DATA(ls).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls-region ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls-product ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls-quantity ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls-revenue ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = lv_row ip_value = ls-sale_date ).
  ADD 1 TO lv_row.
ENDLOOP.
```

### Using bind_table (Recommended)

```abap
lo_worksheet->bind_table(
  ip_table          = lt_sales
  is_table_settings = VALUE #(
    top_left_column  = 'A'
    top_left_row     = 1
    table_style      = zcl_excel_table=>builtinstyle_medium2
    show_row_stripes = abap_true
  )
).
```

## Worksheet Management

```abap
DATA: lo_summary TYPE REF TO zcl_excel_worksheet,
      lo_detail  TYPE REF TO zcl_excel_worksheet.

lo_summary = lo_excel->add_new_worksheet( ).
lo_summary->set_title( 'Summary' ).

lo_detail = lo_excel->add_new_worksheet( ).
lo_detail->set_title( 'Detailed Data' ).

lo_excel->set_active_sheet_index( 1 ).

" Print layout
lo_worksheet->sheet_setup->set_orientation( zcl_excel_sheet_setup=>c_orientation_landscape ).
lo_worksheet->sheet_setup->set_paper_size( zcl_excel_sheet_setup=>c_papersize_a4 ).
lo_worksheet->freeze_panes( ip_num_rows = 1 ip_num_columns = 1 ).
```

## File Output Options

```abap
" Standard Excel 2007+ (.xlsx)
CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
DATA(lv_xlsx) = lo_writer->write_file( lo_excel ).

" CSV - see the CSV Export guide for advanced options (skip hidden rows/cols)
DATA lo_csv TYPE REF TO zcl_excel_writer_csv.
CREATE OBJECT lo_csv.
DATA(lv_csv) = lo_csv->write_file( lo_excel ).

" Huge file writer - memory-efficient for very large datasets
DATA lo_huge TYPE REF TO zcl_excel_writer_huge_file.
CREATE OBJECT lo_huge.
DATA(lv_huge) = lo_huge->write_file( lo_excel ).
```

For CSV options including skipping hidden rows and columns, see **[CSV Export](/guide/csv-export)**.

## Error Handling

```abap
TRY.
    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->add_new_worksheet( ).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Test' ).
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    DATA(lv_result) = lo_writer->write_file( lo_excel ).
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE |Excel error: { lx_excel->get_text( ) }| TYPE 'E'.
ENDTRY.
```

## Next Steps

- **[Reading Excel Files](/guide/reading-excel)** - Read existing Excel files
- **[Formatting](/guide/formatting)** - Add professional styling
- **[Worksheets](/guide/worksheets)** - Multiple sheets, comments, copy semantics
- **[Formulas](/guide/formulas)** - Adding calculations
- **[Data Conversion](/guide/data-conversion)** - Converting ABAP data structures
- **[CSV Export](/guide/csv-export)** - Exporting to CSV with skip-hidden options
- **[ALV Integration](/guide/alv-integration)** - Converting ALV grids (on-premise)
- **[Cloud Compatibility](/guide/cloud-compatibility)** - BTP / S/4HANA Cloud notes
