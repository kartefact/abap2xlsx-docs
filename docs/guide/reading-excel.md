# Reading Excel Files

Comprehensive guide to reading and parsing existing Excel files with abap2xlsx.

## Basic File Reading

### Loading an Excel File

The primary class for reading Excel files is `zcl_excel_reader_2007`, which handles Excel 2007+ (.xlsx) format files.

```abap
" Basic Excel file reading
REPORT zread_excel_basic.

DATA: lo_reader    TYPE REF TO zif_excel_reader,
      lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lv_file_data TYPE xstring.

START-OF-SELECTION.
  " Populate lv_file_data from an upload dialog, file system, BDS, etc.
  " lv_file_data = ... your file loading logic ...

  " Always instantiate through the interface so the reader is swappable
  CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.

  " load_file parses the ZIP/OOXML structure and builds the object model in memory
  TRY.
      lo_excel = lo_reader->load_file( lv_file_data ).
      MESSAGE 'Excel file loaded successfully' TYPE 'S'.
    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |Error loading Excel file: { lx_excel->get_text( ) }| TYPE 'E'.
  ENDTRY.
```

### Accessing Worksheets

```abap
" Get worksheets from loaded Excel file
DATA: lo_worksheets      TYPE REF TO zcl_excel_worksheets,
      lv_worksheet_count TYPE i.

" The worksheets collection mirrors the tab order in Excel
lo_worksheets      = lo_excel->get_worksheets( ).
lv_worksheet_count = lo_worksheets->size( ).

WRITE: / |Excel file contains { lv_worksheet_count } worksheets|.

" get_active_worksheet returns whichever tab was active when the file was saved
lo_worksheet = lo_excel->get_active_worksheet( ).
WRITE: / 'Active worksheet:', lo_worksheet->get_title( ).

" Index is 1-based — sheet at position 1 is the leftmost tab
lo_worksheet = lo_excel->get_worksheet_by_index( 1 ).

" Name lookup is case-sensitive and must match the tab label exactly
lo_worksheet = lo_excel->get_worksheet_by_name( 'Sheet1' ).
```

## Reading Cell Data

### Individual Cell Access

```abap
" Read individual cells
DATA: lv_cell_value   TYPE string,
      lv_cell_formula TYPE string.

" get_cell always returns a string; cast to a typed variable as needed
lv_cell_value = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).
WRITE: / 'Cell A1:', lv_cell_value.

" get_cell_formula returns the formula string (e.g. 'SUM(B2:B10)') or empty
lv_cell_formula = lo_worksheet->get_cell_formula( ip_column = 'B' ip_row = 1 ).
IF lv_cell_formula IS NOT INITIAL.
  WRITE: / 'Cell B1 formula:', lv_cell_formula.
ENDIF.

" A quick existence check — empty string means no cell content was stored
IF lo_worksheet->get_cell( ip_column = 'C' ip_row = 1 ) IS NOT INITIAL.
  WRITE: / 'Cell C1 has content'.
ENDIF.
```

### Reading Cell Ranges

```abap
" Determine the bounding box of the used range
DATA: lv_highest_row       TYPE i,
      lv_highest_col       TYPE i,
      lv_highest_col_alpha TYPE string.

lv_highest_row       = lo_worksheet->get_highest_row( ).
lv_highest_col       = lo_worksheet->get_highest_column( ).

" convert_column2alpha converts a numeric column index to its letter equivalent
lv_highest_col_alpha = zcl_excel_common=>convert_column2alpha( lv_highest_col ).

WRITE: / |Data range: A1:{ lv_highest_col_alpha }{ lv_highest_row }|.

" Iterate every cell in the used range
DATA: lv_row       TYPE i,
      lv_col       TYPE i,
      lv_col_alpha TYPE string.

DO lv_highest_row TIMES.
  lv_row = sy-index.  " sy-index is 1-based inside DO

  DO lv_highest_col TIMES.
    lv_col       = sy-index.
    lv_col_alpha = zcl_excel_common=>convert_column2alpha( lv_col ).

    lv_cell_value = lo_worksheet->get_cell(
      ip_column = lv_col_alpha
      ip_row    = lv_row
    ).

    IF lv_cell_value IS NOT INITIAL.  " Skip truly empty cells
      WRITE: / |{ lv_col_alpha }{ lv_row }: { lv_cell_value }|.
    ENDIF.
  ENDDO.
ENDDO.
```

## Reading Excel Tables

Abap2xlsx supports reading structured Excel tables (ListObjects), including their column definitions and totals-row configurations.

### Accessing Tables in a Worksheet

```abap
" Read tables defined on a worksheet
DATA: lo_tables      TYPE REF TO zcl_excel_worksheet_tables,
      lo_table       TYPE REF TO zcl_excel_table,
      lv_table_count TYPE i.

lo_tables      = lo_worksheet->get_tables( ).
lv_table_count = lo_tables->size( ).

WRITE: / |Worksheet contains { lv_table_count } table(s)|.

DO lv_table_count TIMES.
  lo_table = lo_tables->get( sy-index ).
  WRITE: / 'Table name: ', lo_table->get_name( ).   " The name shown in the Name Box in Excel
  WRITE: / 'Table range:', lo_table->get_ref( ).    " OOXML ref string, e.g. 'A1:D20'
ENDDO.
```

### Reading Table Column Totals Row Functions

> **New in Feb 2026** (PR [#1296](https://github.com/abap2xlsx/abap2xlsx/pull/1296)) — The reader now correctly populates the `totalsRowFunction` attribute on each table column when reading an existing `.xlsx` file. Previously this attribute was lost on read, breaking round-trip fidelity for files with SUM/COUNT/AVERAGE totals rows.

The value of `totalsRowFunction` matches standard OOXML attribute values: `sum`, `count`, `average`, `max`, `min`, `stdDev`, `var`, `countNums`, `custom`.

```abap
" Read table columns and inspect their totals-row configuration
DATA: lo_table_columns TYPE REF TO zcl_excel_table_columns,
      lo_table_col     TYPE REF TO zcl_excel_table_column,
      lv_totals_func   TYPE string,
      lv_col_count     TYPE i.

lo_table_columns = lo_table->get_table_columns( ).
lv_col_count     = lo_table_columns->size( ).

DO lv_col_count TIMES.
  lo_table_col   = lo_table_columns->get( sy-index ).

  " get_totals_row_function returns '' when the column has no totals-row formula
  lv_totals_func = lo_table_col->get_totals_row_function( ).

  IF lv_totals_func IS NOT INITIAL.
    WRITE: / |Column { lo_table_col->get_name( ) } totals: { lv_totals_func }|.
  ENDIF.
ENDDO.
```

## Converting Excel Data to Internal Tables

### Automatic Table Conversion

```abap
" Define a flat structure that mirrors the columns in the Excel sheet
TYPES: BEGIN OF ty_employee,
         emp_id     TYPE i,
         name       TYPE string,
         department TYPE string,
         salary     TYPE p DECIMALS 2,
         hire_date  TYPE d,
       END OF ty_employee.

DATA: lt_employees TYPE TABLE OF ty_employee,
      ls_employee  TYPE ty_employee.

" Row 1 is assumed to hold column headers; data starts at row 2
DATA: lv_data_row TYPE i VALUE 2.

DO lv_highest_row - 1 TIMES.  " -1 because we skip the header row
  CLEAR ls_employee.

  ls_employee-emp_id     = lo_worksheet->get_cell( ip_column = 'A' ip_row = lv_data_row ).
  ls_employee-name       = lo_worksheet->get_cell( ip_column = 'B' ip_row = lv_data_row ).
  ls_employee-department = lo_worksheet->get_cell( ip_column = 'C' ip_row = lv_data_row ).
  ls_employee-salary     = lo_worksheet->get_cell( ip_column = 'D' ip_row = lv_data_row ).
  ls_employee-hire_date  = lo_worksheet->get_cell( ip_column = 'E' ip_row = lv_data_row ).

  " Only append if there is at least an employee ID — skip completely empty rows
  IF ls_employee-emp_id IS NOT INITIAL.
    APPEND ls_employee TO lt_employees.
  ENDIF.

  ADD 1 TO lv_data_row.
ENDDO.

WRITE: / |Imported { lines( lt_employees ) } employee records|.
```

### Dynamic Field Mapping

```abap
" Read the header row to discover column names at runtime
DATA: lt_headers TYPE TABLE OF string,
      lv_header  TYPE string.

DO lv_highest_col TIMES.
  lv_col_alpha = zcl_excel_common=>convert_column2alpha( sy-index ).
  lv_header    = lo_worksheet->get_cell( ip_column = lv_col_alpha ip_row = 1 ).

  IF lv_header IS NOT INITIAL.  " Stop at the first blank header
    APPEND lv_header TO lt_headers.
  ENDIF.
ENDDO.

LOOP AT lt_headers INTO lv_header.
  WRITE: / |Column { sy-tabix }: { lv_header }|.
ENDLOOP.
```

## Handling Different Data Types

### Data Type Conversion

```abap
" Handle different Excel data types
METHOD convert_excel_cell_value.
  DATA: lv_raw_value     TYPE string,
        lv_date_value    TYPE d,
        lv_number_value  TYPE p DECIMALS 2,
        lv_integer_value TYPE i.

  " Always read the raw string first
  lv_raw_value = lo_worksheet->get_cell( ip_column = ip_column ip_row = ip_row ).

  CASE ip_data_type.
    WHEN 'DATE'.
      " excel_string_to_date converts the Excel serial date string to ABAP date type
      lv_date_value      = zcl_excel_common=>excel_string_to_date( lv_raw_value ).
      rv_converted_value = lv_date_value.
    WHEN 'NUMBER'.
      lv_number_value    = lv_raw_value.  " ABAP implicit conversion from string
      rv_converted_value = lv_number_value.
    WHEN 'INTEGER'.
      lv_integer_value   = lv_raw_value.
      rv_converted_value = lv_integer_value.
    WHEN OTHERS.
      rv_converted_value = lv_raw_value.  " Return as-is for string/unknown types
  ENDCASE.
ENDMETHOD.
```

### Handling Formulas and Calculated Values

```abap
" A formula cell stores both the formula text and its last-calculated value
DATA: lv_formula          TYPE string,
      lv_calculated_value TYPE string.

lv_formula          = lo_worksheet->get_cell_formula( ip_column = 'F' ip_row = 10 ).
lv_calculated_value = lo_worksheet->get_cell( ip_column = 'F' ip_row = 10 ).

IF lv_formula IS NOT INITIAL.
  WRITE: / |Cell F10 formula: { lv_formula }|.
  WRITE: / |Calculated value: { lv_calculated_value }|.  " Value cached at last save
ELSE.
  WRITE: / |Cell F10 value: { lv_calculated_value }|.    " Plain value, no formula
ENDIF.
```

## Reading Worksheet Properties

### Worksheet Metadata

```abap
" Get worksheet properties
DATA: lo_sheet_setup TYPE REF TO zcl_excel_sheet_setup.

WRITE: / 'Worksheet title:', lo_worksheet->get_title( ).

lo_sheet_setup = lo_worksheet->get_sheet_setup( ).
IF lo_sheet_setup IS BOUND.
  WRITE: / 'Print orientation:', lo_sheet_setup->get_orientation( ).  " e.g. 'landscape'
  WRITE: / 'Paper size:',        lo_sheet_setup->get_paper_size( ).   " Numeric OOXML code
ENDIF.
```

### Reading Comments and Annotations

```abap
" Read cell comments
DATA: lo_comments TYPE REF TO zcl_excel_comments,
      lo_comment  TYPE REF TO zcl_excel_comment.

" get_comments() returns a *copy* of the comments collection (see note below)
lo_comments = lo_worksheet->get_comments( ).

lo_comment = lo_comments->get_comment( ip_column = 'A' ip_row = 1 ).
IF lo_comment IS BOUND.
  WRITE: / 'Comment on A1:', lo_comment->get_text( ).
ENDIF.
```

> **Note (Jun 2025 — PR [#1317](https://github.com/abap2xlsx/abap2xlsx/pull/1317)):** `get_comments()` now returns a **copy** of the internal comments collection by default. Modifications to the returned object do not affect the worksheet's internal state. If you need to manipulate the live collection, obtain the reference before the copy is taken, or pass the comments instance directly into the worksheet constructor.

## SAP Note 2922674 — XML Namespace Handling

> **New in Nov 2025 (PR [#1349](https://github.com/abap2xlsx/abap2xlsx/pull/1349))** — The XML namespace handling introduced by SAP Note 2922674 was previously only applied in the writer (`render_xml_document`). The reader now also handles files that contain this additional namespace declaration, preventing data loss when round-tripping files produced on certain SAP releases.

No code changes are required in your application. The fix is applied transparently inside `zcl_excel_reader_2007`.

## Accessing `get_style_from_guid` Publicly

> **New in Jun 2025 (PR [#1315](https://github.com/abap2xlsx/abap2xlsx/pull/1315))** — `zcl_excel->get_style_from_guid()` is now a **public** method. Previously it was only accessible internally. You can now look up a style object by its GUID from outside the class:

```abap
DATA: lo_style TYPE REF TO zcl_excel_style,
      lv_guid  TYPE char32.

" Retrieve a style reference by GUID (e.g. obtained from a cell's style attribute)
lo_style = lo_excel->get_style_from_guid( lv_guid ).
IF lo_style IS BOUND.
  " font->get_structure() returns a flat structure with name, size, bold, italic, etc.
  DATA(ls_font) = lo_style->font->get_structure( ).
  WRITE: / 'Font name:', ls_font-name.
ENDIF.
```

This also fixed a code duplication in `zcl_excel_worksheet->check_rtf` (which previously re-implemented the same lookup inline) and corrected a comparison operator bug (`>` instead of `<`) in that method.

## Advanced Reading Features

### Reading Merged Cells

```abap
" Detect merged cell ranges
DATA: lo_ranges TYPE REF TO zcl_excel_ranges,
      lo_range  TYPE REF TO zcl_excel_range.

lo_ranges = lo_worksheet->get_merge( ).

DATA: lv_range_count TYPE i.
lv_range_count = lo_ranges->size( ).

DO lv_range_count TIMES.
  lo_range = lo_ranges->get( sy-index ).
  " get_value() returns the range address string, e.g. 'B2:D4'
  WRITE: / |Merged range { sy-index }: { lo_range->get_value( ) }|.
ENDDO.
```

### Reading Conditional Formatting

```abap
" Read conditional formatting rules (count only — use the rules collection for details)
DATA: lo_cond_formats TYPE REF TO zcl_excel_styles_cond.

lo_cond_formats = lo_worksheet->get_styles_cond( ).
WRITE: / |Worksheet has { lo_cond_formats->size( ) } conditional formatting rules|.
```

## Error Handling and Validation

### Robust File Reading

```abap
METHOD read_excel_file_safely.
  DATA: lo_reader TYPE REF TO zif_excel_reader,
        lo_excel  TYPE REF TO zcl_excel.

  TRY.
      " Sanity-check the payload before attempting a full parse
      IF xstrlen( iv_file_data ) < 100.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING error = 'File too small or empty'.
      ENDIF.

      " All .xlsx files are ZIP archives; check for the PK header signature
      DATA(lv_header) = iv_file_data(4).
      IF lv_header <> '504B0304'.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING error = 'Invalid Excel file format'.
      ENDIF.

      CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
      lo_excel = lo_reader->load_file( iv_file_data ).

      " A valid workbook must have at least one worksheet
      IF lo_excel->get_worksheets( )->size( ) = 0.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING error = 'No worksheets found in file'.
      ENDIF.

      rv_excel = lo_excel.

    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |Excel reading error: { lx_excel->get_text( ) }| TYPE 'E'.
    CATCH cx_root INTO DATA(lx_root).
      " Catch unexpected runtime errors (e.g. memory, XML parse faults)
      MESSAGE |Unexpected error: { lx_root->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

## Performance Considerations

### Efficient Reading Strategies

```abap
METHOD read_excel_efficiently.
  DATA(lv_max_row) = lo_worksheet->get_highest_row( ).
  DATA(lv_max_col) = lo_worksheet->get_highest_column( ).

  DATA: lv_row        TYPE i VALUE 1,
        lv_empty_rows TYPE i VALUE 0.

  DO lv_max_row TIMES.
    DATA(lv_row_empty) = abap_true.

    DO lv_max_col TIMES.
      DATA(lv_col_alpha) = zcl_excel_common=>convert_column2alpha( sy-index ).
      IF lo_worksheet->get_cell( ip_column = lv_col_alpha ip_row = lv_row ) IS NOT INITIAL.
        lv_row_empty = abap_false.
        EXIT.  " Found at least one non-empty cell — no need to check further columns
      ENDIF.
    ENDDO.

    IF lv_row_empty = abap_true.
      ADD 1 TO lv_empty_rows.
    ENDIF.

    ADD 1 TO lv_row.
  ENDDO.

  WRITE: / |Skipped { lv_empty_rows } empty rows|.
ENDMETHOD.
```

## Next Steps

After mastering Excel file reading:

- **[Working with Worksheets](/guide/worksheets)** - Navigate between multiple sheets
- **[Cell Formatting](/guide/formatting)** - Understand and preserve formatting
- **[Data Conversion](/guide/data-conversion)** - Converting Excel data to ABAP structures
- **[Performance Optimization](/guide/performance)** - Efficient reading strategies for large files
- **[Cloud Compatibility](/guide/cloud-compatibility)** - Using abap2xlsx on S/4HANA Cloud and BTP
- **[Changelog](/guide/changelog)** - Full history of recent changes

## Common Reading Patterns

### Complete File Processing Example

```abap
METHOD process_excel_upload.
  DATA: lo_reader         TYPE REF TO zif_excel_reader,
        lo_excel          TYPE REF TO zcl_excel,
        lo_worksheet      TYPE REF TO zcl_excel_worksheet,
        lt_processed_data TYPE TABLE OF your_structure.

  TRY.
      CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
      lo_excel = lo_reader->load_file( iv_file_data ).

      " Always work from the active sheet unless you have a specific sheet name
      lo_worksheet = lo_excel->get_active_worksheet( ).

      " Delegate the actual cell-to-field mapping to a dedicated helper method
      lt_processed_data = convert_worksheet_to_table( lo_worksheet ).

      LOOP AT lt_processed_data INTO DATA(ls_data).
        " Your business logic here
      ENDLOOP.

    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |File processing error: { lx_excel->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```
