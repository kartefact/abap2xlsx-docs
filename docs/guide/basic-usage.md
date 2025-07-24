# Basic Usage

Comprehensive guide to creating your first Excel files with abap2xlsx.

## Core Concepts

### The Excel Object Model

abap2xlsx follows Excel's object hierarchy:

```abap
" Basic object hierarchy
DATA: lo_excel TYPE REF TO zcl_excel,               " Workbook
      lo_worksheet TYPE REF TO zcl_excel_worksheet, " Worksheet
      lo_writer TYPE REF TO zif_excel_writer.       " File writer

" Create workbook (top-level container)
CREATE OBJECT lo_excel.

" Add worksheet to workbook
lo_worksheet = lo_excel->add_new_worksheet( ).

" Write workbook to file
CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
DATA(lv_file) = lo_writer->write_file( lo_excel ).
```

### Understanding Cell References

Excel uses column letters and row numbers for cell addressing:

```abap
" Different ways to reference cells
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Cell A1' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Cell B1' ).
lo_worksheet->set_cell( ip_column = 'AA' ip_row = 1 ip_value = 'Cell AA1' ).

" Reading cell values
DATA(lv_value) = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).
```

## Creating Your First Workbook

### Step-by-Step Workbook Creation

```abap
REPORT zcreate_first_workbook.

DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zif_excel_writer.

START-OF-SELECTION.
  " Step 1: Create workbook
  CREATE OBJECT lo_excel.
  
  " Step 2: Get default worksheet (automatically created)
  lo_worksheet = lo_excel->get_active_worksheet( ).
  
  " Step 3: Set worksheet properties
  lo_worksheet->set_title( 'My Data' ).
  
  " Step 4: Add data to cells
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Product' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Quantity' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Price' ).
  
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Laptop' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 10 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = '999.99' ).
  
  " Step 5: Generate Excel file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_excel_file) = lo_writer->write_file( lo_excel ).
  
  " Step 6: Handle the generated file (download, save, etc.)
  MESSAGE 'Excel file created successfully' TYPE 'S'.
```

## Working with Data Types

### ABAP to Excel Data Type Mapping

```abap
" Different ABAP data types and how they appear in Excel
DATA: lv_string TYPE string VALUE 'Text Value',
      lv_integer TYPE i VALUE 42,
      lv_decimal TYPE p DECIMALS 2 VALUE '123.45',
      lv_date TYPE d VALUE '20231225',
      lv_time TYPE t VALUE '143000',
      lv_boolean TYPE abap_bool VALUE abap_true.

" Add different data types to Excel
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = lv_string ).   " Text
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = lv_integer ).  " Number
lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = lv_decimal ).  " Decimal
lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = lv_date ).     " Date
lo_worksheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = lv_time ).     " Time
lo_worksheet->set_cell( ip_column = 'A' ip_row = 6 ip_value = lv_boolean ).  " Boolean
```

### Handling Special Characters and Unicode

```abap
" Unicode and special characters
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Café' ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Müller' ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = '你好' ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = '€1,234.56' ).

" Line breaks in cells
DATA(lv_multiline) = |Line 1{ cl_abap_char_utilities=>cr_lf }Line 2|.
lo_worksheet->set_cell( ip_column = 'A' ip_row = 5 ip_value = lv_multiline ).
```

## Working with Internal Tables

### Converting Internal Tables to Excel

```abap
" Define structure and internal table
TYPES: BEGIN OF ty_sales_record,
         region TYPE string,
         product TYPE string,
         quantity TYPE i,
         revenue TYPE p DECIMALS 2,
         sale_date TYPE d,
       END OF ty_sales_record.

DATA: lt_sales TYPE TABLE OF ty_sales_record.

" Fill sample data
APPEND VALUE #( region = 'North' product = 'Laptop' quantity = 5 
                revenue = '4999.95' sale_date = '20231201' ) TO lt_sales.
APPEND VALUE #( region = 'South' product = 'Mouse' quantity = 25 
                revenue = '499.75' sale_date = '20231202' ) TO lt_sales.

" Method 1: Manual loop through internal table
DATA: lv_row TYPE i VALUE 2.  " Start from row 2 (row 1 for headers)

" Add headers
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Region' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Product' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Quantity' ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Revenue' ).
lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Sale Date' ).

" Add data rows
LOOP AT lt_sales INTO DATA(ls_sales).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_sales-region ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_sales-product ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_sales-quantity ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_sales-revenue ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = lv_row ip_value = ls_sales-sale_date ).
  ADD 1 TO lv_row.
ENDLOOP.
```

### Using Table Binding (Recommended)

```abap
" Method 2: Use bind_table for automatic conversion
lo_worksheet->bind_table( 
  ip_table = lt_sales
  is_table_settings = VALUE #(
    top_left_column = 'A'
    top_left_row = 1
    table_style = zcl_excel_table=>builtinstyle_medium2
    show_row_stripes = abap_true
    show_first_column = abap_false
    show_last_column = abap_false
  )
).
```

## Cell Operations

### Reading and Writing Cells

```abap
" Writing to cells
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello World' ).

" Reading from cells
DATA(lv_cell_value) = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).

" Check if cell has value
IF lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ) IS NOT INITIAL.
  WRITE: / 'Cell A1 has content'.
ENDIF.

" Clear cell content
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = '' ).
```

### Working with Cell Ranges

```abap
" Set values for a range of cells
DATA: lv_start_row TYPE i VALUE 1,
      lv_end_row TYPE i VALUE 5.

DO lv_end_row TIMES.
  lo_worksheet->set_cell( 
    ip_column = 'A' 
    ip_row = sy-index 
    ip_value = |Row { sy-index }| 
  ).
ENDDO.

" Get worksheet dimensions
DATA: lv_max_row TYPE i,
      lv_max_col TYPE i.

lv_max_row = lo_worksheet->get_highest_row( ).
lv_max_col = lo_worksheet->get_highest_column( ).

WRITE: / |Worksheet contains { lv_max_row } rows and { lv_max_col } columns|.
```

## Worksheet Management

### Creating Multiple Worksheets

```abap
" Create additional worksheets
DATA: lo_summary_sheet TYPE REF TO zcl_excel_worksheet,
      lo_detail_sheet TYPE REF TO zcl_excel_worksheet,
      lo_chart_sheet TYPE REF TO zcl_excel_worksheet.

" Add new worksheets
lo_summary_sheet = lo_excel->add_new_worksheet( ).
lo_summary_sheet->set_title( 'Summary' ).

lo_detail_sheet = lo_excel->add_new_worksheet( ).
lo_detail_sheet->set_title( 'Detailed Data' ).

lo_chart_sheet = lo_excel->add_new_worksheet( ).
lo_chart_sheet->set_title( 'Charts' ).

" Set active worksheet
lo_excel->set_active_sheet_index( 1 ).  " Make first sheet active
```

### Worksheet Properties

```abap
" Configure worksheet properties
lo_worksheet->set_title( 'Sales Report 2023' ).

" Set print properties
lo_worksheet->sheet_setup->set_orientation( zcl_excel_sheet_setup=>c_orientation_landscape ).
lo_worksheet->sheet_setup->set_paper_size( zcl_excel_sheet_setup=>c_papersize_a4 ).

" Set margins (in inches)
lo_worksheet->sheet_setup->set_margin_left( '0.75' ).
lo_worksheet->sheet_setup->set_margin_right( '0.75' ).
lo_worksheet->sheet_setup->set_margin_top( '1.0' ).
lo_worksheet->sheet_setup->set_margin_bottom( '1.0' ).

" Freeze panes
lo_worksheet->freeze_panes( ip_num_rows = 1 ip_num_columns = 1 ).
```

## File Output Options

### Different Writer Types

```abap
" Standard Excel 2007+ (.xlsx) - Most common
CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
DATA(lv_xlsx_file) = lo_writer->write_file( lo_excel ).

" CSV format - For simple data exchange
DATA: lo_csv_writer TYPE REF TO zcl_excel_writer_csv.
CREATE OBJECT lo_csv_writer.
DATA(lv_csv_file) = lo_csv_writer->write_file( lo_excel ).

" Huge file writer - For very large datasets (memory efficient)
DATA: lo_huge_writer TYPE REF TO zcl_excel_writer_huge_file.
CREATE OBJECT lo_huge_writer.
DATA(lv_huge_file) = lo_huge_writer->write_file( lo_excel ).
```

### File Properties and Metadata

```abap
" Set workbook properties
lo_excel->set_properties( 
  ip_title = 'Sales Analysis Report'
  ip_subject = 'Monthly sales data analysis'
  ip_creator = sy-uname
  ip_description = 'Generated by ABAP program'
).
```

## Error Handling Best Practices

### Comprehensive Error Handling

```abap
" Always wrap Excel operations in TRY-CATCH
TRY.
    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->add_new_worksheet( ).
    
    " Your Excel operations
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Test Data' ).
    
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
    DATA(lv_result) = lo_writer->write_file( lo_excel ).
    
    MESSAGE 'Excel file generated successfully' TYPE 'S'.
    
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE |Excel error: { lx_excel->get_text( ) }| TYPE 'E'.
    
  CATCH cx_sy_create_object_error INTO DATA(lx_create).
    MESSAGE |Object creation error: { lx_create->get_text( ) }| TYPE 'E'.
    
  CATCH cx_root INTO DATA(lx_root).
    MESSAGE |Unexpected error: { lx_root->get_text( ) }| TYPE 'E'.
ENDTRY.
```

## Performance Tips

### Efficient Cell Operations

```abap
" Efficient: Set cells in row order
DATA: lv_row TYPE i VALUE 1.
LOOP AT lt_data INTO DATA(ls_data).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_data-field1 ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_data-field2 ).
  ADD 1 TO lv_row.
ENDLOOP.

" Less efficient: Random cell access
" Avoid jumping around between different areas of the worksheet
```

### Memory Management

```abap
" Clear objects when finished
CLEAR: lo_excel, lo_worksheet, lo_writer.

" For large datasets, consider processing in chunks
DATA: lv_chunk_size TYPE i VALUE 1000.
" Process data in batches of 1000 rows
```

## Next Steps

After mastering basic usage:

- **[Reading Excel Files](/guide/reading-excel)** - Learn to read existing Excel files
- **[Cell Formatting](/guide/formatting)** - Add professional styling
- **[Working with Worksheets](/guide/worksheets)** - Multiple sheets, navigation, and organization
- **[Excel Formulas](/guide/formulas)** - Adding calculations and formulas
- **[Data Conversion](/guide/data-conversion)** - Converting ABAP data structures to Excel
- **[ALV Integration](/guide/alv-integration)** - Converting ALV grids to Excel format

## Common Patterns Summary

### Quick Reference for Common Operations

```abap
" Create workbook and worksheet
CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).

" Set cell values
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Data' ).

" Bind internal table
lo_worksheet->bind_table( ip_table = lt_data ).

" Generate file
CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
DATA(lv_file) = lo_writer->write_file( lo_excel ).
```

This covers the fundamental operations you'll need for most Excel generation scenarios. The next guides will dive deeper into specific features and advanced capabilities.
