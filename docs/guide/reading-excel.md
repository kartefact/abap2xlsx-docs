# Reading Excel Files

Comprehensive guide to reading and parsing existing Excel files with abap2xlsx.

## Basic File Reading

### Loading an Excel File

The primary class for reading Excel files is `zcl_excel_reader_2007`, which handles Excel 2007+ (.xlsx) format files .

```abap
" Basic Excel file reading
REPORT zread_excel_basic.

DATA: lo_reader TYPE REF TO zif_excel_reader,
      lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lv_file_data TYPE xstring.

START-OF-SELECTION.
  " Load Excel file data (from upload, file system, etc.)
  " lv_file_data = ... your file loading logic
  
  " Create reader instance
  CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
  
  " Load the Excel file
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
DATA: lo_worksheets TYPE REF TO zcl_excel_worksheets,
      lv_worksheet_count TYPE i.

" Get all worksheets
lo_worksheets = lo_excel->get_worksheets( ).
lv_worksheet_count = lo_worksheets->size( ).

WRITE: / |Excel file contains { lv_worksheet_count } worksheets|.

" Get active worksheet
lo_worksheet = lo_excel->get_active_worksheet( ).
WRITE: / 'Active worksheet:', lo_worksheet->get_title( ).

" Get worksheet by index (1-based)
lo_worksheet = lo_excel->get_worksheet_by_index( 1 ).

" Get worksheet by name
lo_worksheet = lo_excel->get_worksheet_by_name( 'Sheet1' ).
```

## Reading Cell Data

### Individual Cell Access

```abap
" Read individual cells
DATA: lv_cell_value TYPE string,
      lv_cell_formula TYPE string.

" Read cell value
lv_cell_value = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).
WRITE: / 'Cell A1:', lv_cell_value.

" Read cell formula
lv_cell_formula = lo_worksheet->get_cell_formula( ip_column = 'B' ip_row = 1 ).
IF lv_cell_formula IS NOT INITIAL.
  WRITE: / 'Cell B1 formula:', lv_cell_formula.
ENDIF.

" Check if cell exists and has content
IF lo_worksheet->get_cell( ip_column = 'C' ip_row = 1 ) IS NOT INITIAL.
  WRITE: / 'Cell C1 has content'.
ENDIF.
```

### Reading Cell Ranges

```abap
" Get worksheet dimensions
DATA: lv_highest_row TYPE i,
      lv_highest_col TYPE i,
      lv_highest_col_alpha TYPE string.

lv_highest_row = lo_worksheet->get_highest_row( ).
lv_highest_col = lo_worksheet->get_highest_column( ).
lv_highest_col_alpha = zcl_excel_common=>convert_column2alpha( lv_highest_col ).

WRITE: / |Data range: A1:{ lv_highest_col_alpha }{ lv_highest_row }|.

" Read all data in a range
DATA: lv_row TYPE i,
      lv_col TYPE i,
      lv_col_alpha TYPE string.

DO lv_highest_row TIMES.
  lv_row = sy-index.
  
  DO lv_highest_col TIMES.
    lv_col = sy-index.
    lv_col_alpha = zcl_excel_common=>convert_column2alpha( lv_col ).
    
    lv_cell_value = lo_worksheet->get_cell( 
      ip_column = lv_col_alpha 
      ip_row = lv_row 
    ).
    
    IF lv_cell_value IS NOT INITIAL.
      WRITE: / |{ lv_col_alpha }{ lv_row }: { lv_cell_value }|.
    ENDIF.
  ENDDO.
ENDDO.
```

## Converting Excel Data to Internal Tables

### Automatic Table Conversion

```abap
" Define target structure
TYPES: BEGIN OF ty_employee,
         emp_id TYPE i,
         name TYPE string,
         department TYPE string,
         salary TYPE p DECIMALS 2,
         hire_date TYPE d,
       END OF ty_employee.

DATA: lt_employees TYPE TABLE OF ty_employee,
      ls_employee TYPE ty_employee.

" Read data starting from row 2 (assuming row 1 has headers)
DATA: lv_data_row TYPE i VALUE 2.

DO lv_highest_row - 1 TIMES.  " Skip header row
  CLEAR ls_employee.
  
  " Map Excel columns to structure fields
  ls_employee-emp_id = lo_worksheet->get_cell( ip_column = 'A' ip_row = lv_data_row ).
  ls_employee-name = lo_worksheet->get_cell( ip_column = 'B' ip_row = lv_data_row ).
  ls_employee-department = lo_worksheet->get_cell( ip_column = 'C' ip_row = lv_data_row ).
  ls_employee-salary = lo_worksheet->get_cell( ip_column = 'D' ip_row = lv_data_row ).
  ls_employee-hire_date = lo_worksheet->get_cell( ip_column = 'E' ip_row = lv_data_row ).
  
  " Only add if row has data
  IF ls_employee-emp_id IS NOT INITIAL.
    APPEND ls_employee TO lt_employees.
  ENDIF.
  
  ADD 1 TO lv_data_row.
ENDDO.

WRITE: / |Imported { lines( lt_employees ) } employee records|.
```

### Dynamic Field Mapping

```abap
" Read headers dynamically
DATA: lt_headers TYPE TABLE OF string,
      lv_header TYPE string.

" Read header row
DO lv_highest_col TIMES.
  lv_col_alpha = zcl_excel_common=>convert_column2alpha( sy-index ).
  lv_header = lo_worksheet->get_cell( ip_column = lv_col_alpha ip_row = 1 ).
  
  IF lv_header IS NOT INITIAL.
    APPEND lv_header TO lt_headers.
  ENDIF.
ENDDO.

" Display headers
LOOP AT lt_headers INTO lv_header.
  WRITE: / |Column { sy-tabix }: { lv_header }|.
ENDLOOP.
```

## Handling Different Data Types

### Data Type Conversion

```abap
" Handle different Excel data types
METHOD convert_excel_cell_value.
  DATA: lv_raw_value TYPE string,
        lv_date_value TYPE d,
        lv_number_value TYPE p DECIMALS 2,
        lv_integer_value TYPE i.

  lv_raw_value = lo_worksheet->get_cell( ip_column = ip_column ip_row = ip_row ).
  
  " Convert based on expected data type
  CASE ip_data_type.
    WHEN 'DATE'.
      " Excel dates are stored as numbers (days since 1900-01-01)
      lv_date_value = zcl_excel_common=>excel_string_to_date( lv_raw_value ).
      rv_converted_value = lv_date_value.
      
    WHEN 'NUMBER'.
      lv_number_value = lv_raw_value.
      rv_converted_value = lv_number_value.
      
    WHEN 'INTEGER'.
      lv_integer_value = lv_raw_value.
      rv_converted_value = lv_integer_value.
      
    WHEN OTHERS.
      " Keep as string
      rv_converted_value = lv_raw_value.
  ENDCASE.
ENDMETHOD.
```

### Handling Formulas and Calculated Values

```abap
" Read both formula and calculated value
DATA: lv_formula TYPE string,
      lv_calculated_value TYPE string.

lv_formula = lo_worksheet->get_cell_formula( ip_column = 'F' ip_row = 10 ).
lv_calculated_value = lo_worksheet->get_cell( ip_column = 'F' ip_row = 10 ).

IF lv_formula IS NOT INITIAL.
  WRITE: / |Cell F10 formula: { lv_formula }|.
  WRITE: / |Calculated value: { lv_calculated_value }|.
ELSE.
  WRITE: / |Cell F10 value: { lv_calculated_value }|.
ENDIF.
```

## Reading Worksheet Properties

### Worksheet Metadata

```abap
" Get worksheet properties
DATA: lv_sheet_title TYPE string,
      lv_sheet_state TYPE string,
      lo_sheet_setup TYPE REF TO zcl_excel_sheet_setup.

lv_sheet_title = lo_worksheet->get_title( ).
WRITE: / 'Worksheet title:', lv_sheet_title.

" Get print setup information
lo_sheet_setup = lo_worksheet->get_sheet_setup( ).
IF lo_sheet_setup IS BOUND.
  DATA(lv_orientation) = lo_sheet_setup->get_orientation( ).
  DATA(lv_paper_size) = lo_sheet_setup->get_paper_size( ).
  
  WRITE: / 'Print orientation:', lv_orientation.
  WRITE: / 'Paper size:', lv_paper_size.
ENDIF.
```

### Reading Comments and Annotations

```abap
" Read cell comments
DATA: lo_comments TYPE REF TO zcl_excel_comments,
      lo_comment TYPE REF TO zcl_excel_comment.

lo_comments = lo_worksheet->get_comments( ).

" Check if specific cell has a comment
lo_comment = lo_comments->get_comment( ip_column = 'A' ip_row = 1 ).
IF lo_comment IS BOUND.
  DATA(lv_comment_text) = lo_comment->get_text( ).
  WRITE: / 'Comment on A1:', lv_comment_text.
ENDIF.
```

## Advanced Reading Features

### Reading Merged Cells

```abap
" Detect merged cell ranges
DATA: lo_ranges TYPE REF TO zcl_excel_ranges,
      lo_range TYPE REF TO zcl_excel_range.

lo_ranges = lo_worksheet->get_merge( ).

" Iterate through merged ranges
DATA: lv_range_count TYPE i.
lv_range_count = lo_ranges->size( ).

DO lv_range_count TIMES.
  lo_range = lo_ranges->get( sy-index ).
  DATA(lv_range_value) = lo_range->get_value( ).
  
  WRITE: / |Merged range { sy-index }: { lv_range_value }|.
ENDDO.
```

### Reading Conditional Formatting

```abap
" Read conditional formatting rules
DATA: lo_cond_formats TYPE REF TO zcl_excel_styles_cond,
      lo_cond_format TYPE REF TO zcl_excel_style_cond.

lo_cond_formats = lo_worksheet->get_styles_cond( ).

" Process conditional formatting rules
DATA: lv_cond_count TYPE i.
lv_cond_count = lo_cond_formats->size( ).

WRITE: / |Worksheet has { lv_cond_count } conditional formatting rules|.
```

## Error Handling and Validation

### Robust File Reading

```abap
" Comprehensive error handling for file reading
METHOD read_excel_file_safely.
  DATA: lo_reader TYPE REF TO zif_excel_reader,
        lo_excel TYPE REF TO zcl_excel.

  TRY.
      " Validate file format
      IF xstrlen( iv_file_data ) < 100.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING error = 'File too small or empty'.
      ENDIF.
      
      " Check file signature (ZIP format for .xlsx)
      DATA(lv_header) = iv_file_data(4).
      IF lv_header <> '504B0304'.  " ZIP file signature
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING error = 'Invalid Excel file format'.
      ENDIF.
      
      " Create reader and load file
      CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
      lo_excel = lo_reader->load_file( iv_file_data ).
      
      " Validate loaded content
      IF lo_excel->get_worksheets( )->size( ) = 0.
        RAISE EXCEPTION TYPE zcx_excel
          EXPORTING error = 'No worksheets found in file'.
      ENDIF.
      
      rv_excel = lo_excel.
      
    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |Excel reading error: { lx_excel->get_text( ) }| TYPE 'E'.
      
    CATCH cx_root INTO DATA(lx_root).
      MESSAGE |Unexpected error: { lx_root->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

## Performance Considerations

### Efficient Reading Strategies

```abap
" Read only necessary data
METHOD read_excel_efficiently.
  " 1. Check worksheet dimensions first
  DATA(lv_max_row) = lo_worksheet->get_highest_row( ).
  DATA(lv_max_col) = lo_worksheet->get_highest_column( ).
  
  " 2. Skip empty rows/columns
  DATA: lv_row TYPE i VALUE 1,
        lv_empty_rows TYPE i VALUE 0.
  
  DO lv_max_row TIMES.
    " Check if entire row is empty
    DATA(lv_row_empty) = abap_true.
    
    DO lv_max_col TIMES.
      DATA(lv_col_alpha) = zcl_excel_common=>convert_column2alpha( sy-index ).
      IF lo_worksheet->get_cell( ip_column = lv_col_alpha ip_row = lv_row ) IS NOT INITIAL.
        lv_row_empty = abap_false.
        EXIT.
      ENDIF.
    ENDDO.
    
    IF lv_row_empty = abap_true.
      ADD 1 TO lv_empty_rows.
    ELSE.
      " Process non-empty row
      " Your row processing logic here
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

## Common Reading Patterns

### Complete File Processing Example

```abap
" Complete example: Read Excel file and process data
METHOD process_excel_upload.
  DATA: lo_reader TYPE REF TO zif_excel_reader,
        lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lt_processed_data TYPE TABLE OF your_structure.

  TRY.
      " Load and validate file
      CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
      lo_excel = lo_reader->load_file( iv_file_data ).
      
      " Get first worksheet
      lo_worksheet = lo_excel->get_active_worksheet( ).
      
      " Convert to internal table
      lt_processed_data = convert_worksheet_to_table( lo_worksheet ).
      
      " Process the data
      LOOP AT lt_processed_data INTO DATA(ls_data).
        " Your business logic here
      ENDLOOP.
      
    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |File processing error: { lx_excel->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

This guide covers the essential techniques for reading Excel files with abap2xlsx. The reader classes provide comprehensive support for extracting data, formulas, and formatting from Excel files.
