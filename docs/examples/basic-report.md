# Simple Report Generation

Basic examples for creating straightforward Excel reports with abap2xlsx.

## Hello World Example

### Minimal Excel Generation

```abap
REPORT zhello_excel.

DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zif_excel_writer,
      lv_file TYPE xstring.

START-OF-SELECTION.
  " Create Excel workbook
  CREATE OBJECT lo_excel.
  
  " Add worksheet
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Hello World' ).
  
  " Add content
  lo_worksheet->set_cell( 
    ip_column = 'A' 
    ip_row = 1 
    ip_value = 'Hello, World!' 
  ).
  
  " Write to file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  lv_file = lo_writer->write_file( lo_excel ).
  
  " Download or save file
  " Implementation depends on your environment
```

### Basic Data Table Report

```abap
" Simple table report
REPORT zsimple_table_report.

TYPES: BEGIN OF ty_employee,
         id TYPE i,
         name TYPE string,
         department TYPE string,
         salary TYPE p DECIMALS 2,
       END OF ty_employee,
       tt_employees TYPE TABLE OF ty_employee.

DATA: lt_employees TYPE tt_employees,
      lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

START-OF-SELECTION.
  " Fill sample data
  APPEND VALUE #( id = 1 name = 'John Doe' department = 'IT' salary = '75000.00' ) TO lt_employees.
  APPEND VALUE #( id = 2 name = 'Jane Smith' department = 'HR' salary = '65000.00' ) TO lt_employees.
  APPEND VALUE #( id = 3 name = 'Bob Johnson' department = 'Finance' salary = '80000.00' ) TO lt_employees.

  " Create Excel
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Employee Report' ).

  " Add headers
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'ID' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Name' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Department' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Salary' ).

  " Add data rows
  DATA: lv_row TYPE i VALUE 2.
  LOOP AT lt_employees INTO DATA(ls_employee).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_employee-id ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_employee-name ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_employee-department ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_employee-salary ).
    ADD 1 TO lv_row.
  ENDLOOP.

  " Generate file
  DATA: lo_writer TYPE REF TO zif_excel_writer.
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
```

## Adding Basic Formatting

### Header Styling

```abap
" Add professional header formatting
DATA: lo_header_style TYPE REF TO zcl_excel_style.

" Create header style
lo_header_style = lo_excel->add_new_style( ).
lo_header_style->font->bold = abap_true.
lo_header_style->font->color->set_rgb( 'FFFFFF' ).
lo_header_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_header_style->fill->fgcolor->set_rgb( '366092' ).
lo_header_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.

" Apply to header row
lo_worksheet->set_cell( 
  ip_column = 'A' 
  ip_row = 1 
  ip_value = 'ID' 
  ip_style = lo_header_style 
).
" Repeat for other header cells...
```

### Number Formatting

```abap
" Format salary column as currency
DATA: lo_currency_style TYPE REF TO zcl_excel_style.

lo_currency_style = lo_excel->add_new_style( ).
lo_currency_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.

" Apply currency formatting to salary column
LOOP AT lt_employees INTO DATA(ls_employee).
  lo_worksheet->set_cell( 
    ip_column = 'D' 
    ip_row = lv_row 
    ip_value = ls_employee-salary 
    ip_style = lo_currency_style 
  ).
  ADD 1 TO lv_row.
ENDLOOP.
```

### Date Formatting

```abap
" Format date columns properly
DATA: lo_date_style TYPE REF TO zcl_excel_style,
      lv_today TYPE d.

lo_date_style = lo_excel->add_new_style( ).
lo_date_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_ddmmyyyy_new.

" Add date column
lv_today = sy-datum.
lo_worksheet->set_cell( 
  ip_column = 'E' 
  ip_row = 1 
  ip_value = 'Hire Date' 
  ip_style = lo_header_style 
).

lo_worksheet->set_cell( 
  ip_column = 'E' 
  ip_row = 2 
  ip_value = lv_today 
  ip_style = lo_date_style 
).
```

## Adding Visual Enhancements

### Borders and Grid Lines

```abap
" Add borders to data range
DATA: lo_border_style TYPE REF TO zcl_excel_style.

lo_border_style = lo_excel->add_new_style( ).

" Create border objects
CREATE OBJECT lo_border_style->borders->allborders.
lo_border_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
lo_border_style->borders->allborders->border_color->set_rgb( '000000' ).

" Apply borders to data range
lo_worksheet->set_cell_style( 
  ip_range = 'A1:D10'
  ip_style = lo_border_style
).
```

### Column Width Adjustment

```abap
" Auto-adjust column widths for better readability
DATA: lo_column TYPE REF TO zcl_excel_column.

" Set specific column widths
lo_column = lo_worksheet->get_column( 'A' ).
lo_column->set_width( 10 ).

lo_column = lo_worksheet->get_column( 'B' ).
lo_column->set_width( 20 ).

lo_column = lo_worksheet->get_column( 'C' ).
lo_column->set_width( 15 ).

lo_column = lo_worksheet->get_column( 'D' ).
lo_column->set_width( 12 ).
```

## Complete Basic Report Example

```abap
" Complete formatted report example
REPORT zcomplete_basic_report.

TYPES: BEGIN OF ty_sales_data,
         region TYPE string,
         salesperson TYPE string,
         amount TYPE p DECIMALS 2,
         date TYPE d,
       END OF ty_sales_data,
       tt_sales_data TYPE TABLE OF ty_sales_data.

DATA: lt_sales_data TYPE tt_sales_data,
      lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zif_excel_writer.

START-OF-SELECTION.
  " Sample data
  APPEND VALUE #( region = 'North' salesperson = 'Alice' amount = '15000.00' date = '20231201' ) TO lt_sales_data.
  APPEND VALUE #( region = 'South' salesperson = 'Bob' amount = '12500.00' date = '20231202' ) TO lt_sales_data.
  APPEND VALUE #( region = 'East' salesperson = 'Carol' amount = '18000.00' date = '20231203' ) TO lt_sales_data.

  " Create Excel with formatting
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Report' ).

  " Create styles
  DATA: lo_header_style TYPE REF TO zcl_excel_style,
        lo_currency_style TYPE REF TO zcl_excel_style,
        lo_date_style TYPE REF TO zcl_excel_style.

  " Header style
  lo_header_style = lo_excel->add_new_style( ).
  lo_header_style->font->bold = abap_true.
  lo_header_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_header_style->fill->fgcolor->set_rgb( 'CCCCCC' ).

  " Currency style
  lo_currency_style = lo_excel->add_new_style( ).
  lo_currency_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.

  " Date style
  lo_date_style = lo_excel->add_new_style( ).
  lo_date_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_ddmmyyyy_new.

  " Add headers
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Region' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Salesperson' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Amount' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Date' ip_style = lo_header_style ).

  " Add data with formatting
  DATA: lv_row TYPE i VALUE 2.
  LOOP AT lt_sales_data INTO DATA(ls_sales).
    lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_sales-region ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_sales-salesperson ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_sales-amount ip_style = lo_currency_style ).
    lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_sales-date ip_style = lo_date_style ).
    ADD 1 TO lv_row.
  ENDLOOP.

  " Add total row
  lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = 'TOTAL' ip_style = lo_header_style ).
  lo_worksheet->set_cell_formula( 
    ip_column = 'C' 
    ip_row = lv_row 
    ip_formula = |SUM(C2:C{ lv_row - 1 })|
    ip_style = lo_currency_style
  ).

  " Generate and save file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
  
  WRITE: / 'Report generated successfully'.
```

## Best Practices for Basic Reports

1. **Consistent Formatting**: Apply consistent styles throughout the report
2. **Appropriate Data Types**: Use proper number formats for different data types
3. **Clear Headers**: Make headers visually distinct from data
4. **Readable Layout**: Adjust column widths and add borders for clarity
5. **Error Handling**: Always include proper exception handling in production code
