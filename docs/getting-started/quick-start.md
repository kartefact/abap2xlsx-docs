# Quick Start Guide

Get up and running with abap2xlsx in minutes with these essential examples.

## Your First Excel File

### Basic "Hello World" Example

```abap
REPORT zhello_abap2xlsx.

DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zif_excel_writer,
      lv_file TYPE xstring.

START-OF-SELECTION.
  " Create workbook
  CREATE OBJECT lo_excel.
  
  " Add worksheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'My First Sheet' ).
  
  " Add some data
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'World!' ).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Created on' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = sy-datum ).
  
  " Generate Excel file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  lv_file = lo_writer->write_file( lo_excel ).
  
  " File is now ready for download/save
  MESSAGE 'Excel file generated successfully!' TYPE 'S'.
```

## Working with Data Tables

### Converting Internal Table to Excel

```abap
" Define data structure
TYPES: BEGIN OF ty_employee,
         emp_id TYPE i,
         name TYPE string,
         department TYPE string,
         salary TYPE p DECIMALS 2,
         hire_date TYPE d,
       END OF ty_employee.

DATA: lt_employees TYPE TABLE OF ty_employee,
      lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

" Fill sample data
APPEND VALUE #( emp_id = 1001 name = 'John Smith' department = 'IT' 
                salary = '75000.00' hire_date = '20220315' ) TO lt_employees.
APPEND VALUE #( emp_id = 1002 name = 'Jane Doe' department = 'HR' 
                salary = '68000.00' hire_date = '20210820' ) TO lt_employees.
APPEND VALUE #( emp_id = 1003 name = 'Mike Johnson' department = 'Finance' 
                salary = '82000.00' hire_date = '20230110' ) TO lt_employees.

" Create Excel with table binding
CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).
lo_worksheet->set_title( 'Employee List' ).

" Bind internal table to worksheet
lo_worksheet->bind_table( 
  ip_table = lt_employees
  is_table_settings = VALUE #(
    top_left_column = 'A'
    top_left_row = 1
    table_style = zcl_excel_table=>builtinstyle_medium9
    show_row_stripes = abap_true
  )
).
```

## Adding Basic Formatting

### Styled Headers and Data

```abap
" Create styles
DATA: lo_header_style TYPE REF TO zcl_excel_style,
      lo_currency_style TYPE REF TO zcl_excel_style,
      lo_date_style TYPE REF TO zcl_excel_style.

" Header style - bold white text on blue background
lo_header_style = lo_excel->add_new_style( ).
lo_header_style->font->bold = abap_true.
lo_header_style->font->color->set_rgb( 'FFFFFF' ).
lo_header_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_header_style->fill->fgcolor->set_rgb( '4472C4' ).
lo_header_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.

" Currency formatting
lo_currency_style = lo_excel->add_new_style( ).
lo_currency_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.

" Date formatting
lo_date_style = lo_excel->add_new_style( ).
lo_date_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_ddmmyyyy_new.

" Apply styles to specific cells
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Employee ID' ip_style = lo_header_style ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Name' ip_style = lo_header_style ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Department' ip_style = lo_header_style ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Salary' ip_style = lo_header_style ).
lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Hire Date' ip_style = lo_header_style ).
```

## Reading Excel Files

### Basic File Reading

```abap
" Read existing Excel file
DATA: lo_reader TYPE REF TO zif_excel_reader,
      lo_excel_read TYPE REF TO zcl_excel,
      lo_worksheet_read TYPE REF TO zcl_excel_worksheet,
      lv_value TYPE string.

" Load file (lv_file_data contains the Excel file as xstring)
CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
lo_excel_read = lo_reader->load_file( lv_file_data ).

" Get first worksheet
lo_worksheet_read = lo_excel_read->get_active_worksheet( ).

" Read specific cells
lv_value = lo_worksheet_read->get_cell( ip_column = 'A' ip_row = 1 ).
WRITE: / 'Cell A1 contains:', lv_value.

" Get worksheet dimensions
DATA: lv_max_row TYPE i,
      lv_max_col TYPE i.

lv_max_row = lo_worksheet_read->get_highest_row( ).
lv_max_col = lo_worksheet_read->get_highest_column( ).

WRITE: / 'Worksheet has', lv_max_row, 'rows and', lv_max_col, 'columns'.
```

## Working with Multiple Worksheets

### Creating Multiple Sheets

```abap
DATA: lo_sheet1 TYPE REF TO zcl_excel_worksheet,
      lo_sheet2 TYPE REF TO zcl_excel_worksheet,
      lo_sheet3 TYPE REF TO zcl_excel_worksheet.

CREATE OBJECT lo_excel.

" Create multiple worksheets
lo_sheet1 = lo_excel->add_new_worksheet( ).
lo_sheet1->set_title( 'Summary' ).

lo_sheet2 = lo_excel->add_new_worksheet( ).
lo_sheet2->set_title( 'Details' ).

lo_sheet3 = lo_excel->add_new_worksheet( ).
lo_sheet3->set_title( 'Charts' ).

" Add data to different sheets
lo_sheet1->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Summary Report' ).
lo_sheet2->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Detailed Data' ).
lo_sheet3->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Chart Analysis' ).

" Set active sheet
lo_excel->set_active_sheet_index( 1 ).  " Make Summary the active sheet
```

## Adding Formulas

### Basic Excel Formulas

```abap
" Add some numbers and calculate sum
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 100 ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 200 ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 300 ).

" Add SUM formula
lo_worksheet->set_cell_formula( 
  ip_column = 'A' 
  ip_row = 4 
  ip_formula = 'SUM(A1:A3)' 
).

" Add label for the sum
lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Total' ).

" More complex formula with IF condition
lo_worksheet->set_cell_formula(
  ip_column = 'C'
  ip_row = 4
  ip_formula = 'IF(A4>500,"High","Low")'
).
```

## File Output Options

### Different Writer Types

```abap
" Standard Excel 2007+ format (.xlsx)
CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
lv_file = lo_writer->write_file( lo_excel ).

" For very large files (memory efficient)
DATA: lo_huge_writer TYPE REF TO zcl_excel_writer_huge_file.
CREATE OBJECT lo_huge_writer.
lv_file = lo_huge_writer->write_file( lo_excel ).

" CSV format
DATA: lo_csv_writer TYPE REF TO zcl_excel_writer_csv.
CREATE OBJECT lo_csv_writer.
lv_file = lo_csv_writer->write_file( lo_excel ).

" Macro-enabled Excel (.xlsm)
DATA: lo_xlsm_writer TYPE REF TO zcl_excel_writer_xlsm.
CREATE OBJECT lo_xlsm_writer.
lv_file = lo_xlsm_writer->write_file( lo_excel ).
```

## Error Handling

### Basic Exception Handling

```abap
TRY.
    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->add_new_worksheet( ).
    
    " Your Excel operations here
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Safe Operation' ).
    
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    lv_file = lo_writer->write_file( lo_excel ).
    
    MESSAGE 'Excel file created successfully!' TYPE 'S'.
    
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE |Excel error: { lx_excel->get_text( ) }| TYPE 'E'.
    
  CATCH cx_root INTO DATA(lx_root).
    MESSAGE |Unexpected error: { lx_root->get_text( ) }| TYPE 'E'.
ENDTRY.
```

## Next Steps

Now that you've created your first Excel files, explore these topics:

- **[Basic Usage Guide](/guide/basic-usage)** - Detailed explanations of core concepts
- **[Cell Formatting](/guide/formatting)** - Advanced styling and formatting options
- **[Working with Worksheets](/guide/worksheets)** - Multiple sheets, navigation, and organization
- **[Data Conversion](/guide/data-conversion)** - Converting ABAP data structures to Excel
- **[ALV Integration](/guide/alv-integration)** - Converting ALV grids to Excel format

## Common Patterns

### Quick Data Export Pattern

```abap
" Standard pattern for exporting internal table to Excel
METHOD export_to_excel.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zif_excel_writer.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( iv_sheet_name ).
  
  lo_worksheet->bind_table( ip_table = it_data ).
  
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  rv_file = lo_writer->write_file( lo_excel ).
ENDMETHOD.
```

This quick start guide covers the essential patterns you'll use in 90% of abap2xlsx scenarios. For more advanced features, continue to the detailed guides in the next sections.
