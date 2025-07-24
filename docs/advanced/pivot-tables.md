# Pivot Tables

Advanced guide to creating and managing pivot tables with abap2xlsx for dynamic data analysis.

## Understanding Pivot Tables

Pivot tables provide powerful data analysis capabilities by allowing users to summarize, analyze, and present large datasets in a flexible, interactive format. The abap2xlsx library supports creating pivot tables through the table system <cite>src/zcl_excel_table.clas.abap:26</cite>.

## Basic Pivot Table Creation

### Setting Up Data Source

```abap
" Create pivot table from data source
METHOD create_basic_pivot_table.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_pivot_table TYPE REF TO zcl_excel_table.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Data' ).

  " Add source data
  populate_sales_data( lo_worksheet ).

  " Create pivot table
  lo_pivot_table = lo_excel->add_new_table( ).
  lo_pivot_table->set_name( 'SalesPivot' ).
  
  " Configure pivot table settings
  DATA(ls_table_settings) = VALUE zexcel_s_table_settings(
    top_left_column = 'A'
    top_left_row = 1
    table_style = zcl_excel_table=>builtinstyle_pivot_light16
    show_row_stripes = abap_true
  ).
  
  lo_pivot_table->settings = ls_table_settings.
ENDMETHOD.
```

### Pivot Table Structure

```abap
" Define pivot table structure with fields
METHOD define_pivot_structure.
  DATA: lt_field_catalog TYPE zexcel_t_fieldcatalog,
        ls_field_catalog TYPE zexcel_s_fieldcatalog.

  " Row fields
  ls_field_catalog-fieldname = 'REGION'.
  ls_field_catalog-position = 1.
  ls_field_catalog-pivot_area = 'ROW'.
  APPEND ls_field_catalog TO lt_field_catalog.

  ls_field_catalog-fieldname = 'PRODUCT'.
  ls_field_catalog-position = 2.
  ls_field_catalog-pivot_area = 'ROW'.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Column fields
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'QUARTER'.
  ls_field_catalog-position = 3.
  ls_field_catalog-pivot_area = 'COLUMN'.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Value fields with aggregation
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'SALES_AMOUNT'.
  ls_field_catalog-position = 4.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_sum.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Apply field catalog to pivot table
  lo_pivot_table->fieldcat = lt_field_catalog.
ENDMETHOD.
```

## Advanced Pivot Table Features

### Multiple Value Fields

```abap
" Create pivot table with multiple aggregated values
METHOD create_multi_value_pivot.
  DATA: lt_field_catalog TYPE zexcel_t_fieldcatalog,
        ls_field_catalog TYPE zexcel_s_fieldcatalog.

  " Sales Amount - Sum
  ls_field_catalog-fieldname = 'SALES_AMOUNT'.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_sum.
  ls_field_catalog-position = 1.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Sales Amount - Average
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'SALES_AMOUNT'.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_average.
  ls_field_catalog-position = 2.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Quantity - Count
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'QUANTITY'.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_count.
  ls_field_catalog-position = 3.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Unit Price - Max/Min
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'UNIT_PRICE'.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_max.
  ls_field_catalog-position = 4.
  APPEND ls_field_catalog TO lt_field_catalog.

  lo_pivot_table->fieldcat = lt_field_catalog.
ENDMETHOD.
```

### Calculated Fields

```abap
" Add calculated fields to pivot table
METHOD add_calculated_fields.
  DATA: ls_field_catalog TYPE zexcel_s_fieldcatalog.

  " Profit Margin = (Sales - Cost) / Sales
  ls_field_catalog-fieldname = 'PROFIT_MARGIN'.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_custom.
  ls_field_catalog-formula = '(SALES_AMOUNT-COST_AMOUNT)/SALES_AMOUNT'.
  ls_field_catalog-position = 5.
  APPEND ls_field_catalog TO lo_pivot_table->fieldcat.

  " Growth Rate = (Current - Previous) / Previous
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'GROWTH_RATE'.
  ls_field_catalog-pivot_area = 'DATA'.
  ls_field_catalog-totals_function = zcl_excel_table=>totals_function_custom.
  ls_field_catalog-formula = '(CURRENT_SALES-PREVIOUS_SALES)/PREVIOUS_SALES'.
  ls_field_catalog-position = 6.
  APPEND ls_field_catalog TO lo_pivot_table->fieldcat.
ENDMETHOD.
```

## Pivot Table Formatting

### Styling Pivot Tables

```abap
" Apply formatting to pivot table
METHOD format_pivot_table.
  " Use built-in pivot table styles
  DATA(ls_settings) = lo_pivot_table->settings.
  ls_settings-table_style = zcl_excel_table=>builtinstyle_pivot_light16.
  ls_settings-show_row_stripes = abap_true.
  ls_settings-show_column_stripes = abap_false.
  ls_settings-show_first_column = abap_true.
  ls_settings-show_last_column = abap_false.
  
  lo_pivot_table->settings = ls_settings.

  " Apply conditional formatting to data area
  apply_pivot_conditional_formatting( lo_pivot_table ).
ENDMETHOD.

METHOD apply_pivot_conditional_formatting.
  DATA: lo_cond_format TYPE REF TO zcl_excel_style_cond.

  " Highlight high values in data area
  lo_cond_format = lo_worksheet->add_new_style_cond( ).
  lo_cond_format->set_range( 'C3:Z100' ).  " Data area range
  lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_top10 ).
  lo_cond_format->set_rank( 10 ).
  lo_cond_format->set_percent( abap_true ).
  lo_cond_format->set_fill_color( '00FF00' ).
ENDMETHOD.
```

## Data Source Management

### Dynamic Data Ranges

```abap
" Create pivot table with dynamic data source
METHOD create_dynamic_pivot_source.
  DATA: lv_last_row TYPE i,
        lv_last_col TYPE i,
        lv_data_range TYPE string.

  " Determine dynamic data range
  lv_last_row = lo_worksheet->get_highest_row( ).
  lv_last_col = lo_worksheet->get_highest_column( ).
  
  lv_data_range = |A1:{ zcl_excel_common=>convert_column2alpha( lv_last_col ) }{ lv_last_row }|.

  " Set data source range
  lo_pivot_table->settings-data_range = lv_data_range.
  
  " Enable auto-refresh when data changes
  lo_pivot_table->settings-refresh_on_load = abap_true.
ENDMETHOD.
```

### External Data Sources

```abap
" Connect pivot table to external data source
METHOD connect_external_data_source.
  DATA: ls_connection TYPE zexcel_s_data_connection.

  " Configure external data connection
  ls_connection-connection_type = 'DATABASE'.
  ls_connection-server = 'SAP_SERVER'.
  ls_connection-database = 'SALES_DB'.
  ls_connection-query = 'SELECT * FROM SALES_DATA WHERE YEAR = 2023'.
  
  " Apply connection to pivot table
  lo_pivot_table->set_data_connection( ls_connection ).
ENDMETHOD.
```

## Pivot Table Totals and Subtotals

### Configuring Totals

The pivot table system leverages the table totals functionality <cite>src/zcl_excel_table.clas.abap:78-88</cite>:

```abap
" Configure totals and subtotals
METHOD configure_pivot_totals.
  " Check if pivot table has totals configured
  IF lo_pivot_table->has_totals( ) = abap_true.
    " Get totals formula for specific column
    DATA(lv_formula) = lo_pivot_table->get_totals_formula(
      ip_column = 'SALES_AMOUNT'
      ip_function = zcl_excel_table=>totals_function_sum
    ).
    
    MESSAGE |Totals formula: { lv_formula }| TYPE 'I'.
  ENDIF.

  " Configure subtotals for row groups
  LOOP AT lo_pivot_table->fieldcat INTO DATA(ls_field) WHERE pivot_area = 'ROW'.
    ls_field-show_subtotals = abap_true.
    ls_field-subtotal_function = zcl_excel_table=>totals_function_sum.
    MODIFY lo_pivot_table->fieldcat FROM ls_field.
  ENDLOOP.
ENDMETHOD.
```

## Pivot Table Filters

### Page Filters

```abap
" Add page filters to pivot table
METHOD add_page_filters.
  DATA: ls_field_catalog TYPE zexcel_s_fieldcatalog.

  " Year filter
  ls_field_catalog-fieldname = 'YEAR'.
  ls_field_catalog-pivot_area = 'PAGE'.
  ls_field_catalog-position = 1.
  ls_field_catalog-filter_value = '2023'.
  APPEND ls_field_catalog TO lo_pivot_table->fieldcat.

  " Department filter
  CLEAR ls_field_catalog.
  ls_field_catalog-fieldname = 'DEPARTMENT'.
  ls_field_catalog-pivot_area = 'PAGE'.
  ls_field_catalog-position = 2.
  ls_field_catalog-filter_values = VALUE #( ( 'SALES' ) ( 'MARKETING' ) ).
  APPEND ls_field_catalog TO lo_pivot_table->fieldcat.
ENDMETHOD.
```

### Row and Column Filters

```abap
" Configure row and column filters
METHOD configure_field_filters.
  " Filter specific values in row fields
  LOOP AT lo_pivot_table->fieldcat INTO DATA(ls_field) WHERE pivot_area = 'ROW'.
    CASE ls_field-fieldname.
      WHEN 'REGION'.
        ls_field-filter_values = VALUE #( ( 'NORTH' ) ( 'SOUTH' ) ( 'EAST' ) ).
      WHEN 'PRODUCT'.
        ls_field-filter_condition = 'BEGINS_WITH'.
        ls_field-filter_value = 'LAPTOP'.
    ENDCASE.
    MODIFY lo_pivot_table->fieldcat FROM ls_field.
  ENDLOOP.
ENDMETHOD.
```

## Complete Pivot Table Example

### Sales Analysis Dashboard

```abap
" Complete example: Sales analysis pivot table
METHOD create_sales_analysis_pivot.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_pivot_table TYPE REF TO zcl_excel_table,
        lo_writer TYPE REF TO zcl_excel_writer_2007.

  " Initialize workbook
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Analysis' ).

  " Load and prepare data
  load_sales_data( lo_worksheet ).

  " Create pivot table
  lo_pivot_table = lo_excel->add_new_table( ).
  lo_pivot_table->set_name( 'SalesAnalysisPivot' ).

  " Configure pivot structure
  configure_sales_pivot_structure( lo_pivot_table ).
  
  " Apply formatting and styling
  format_sales_pivot_table( lo_pivot_table ).
  
  " Add calculated fields
  add_sales_calculated_fields( lo_pivot_table ).
  
  " Configure filters
  setup_sales_filters( lo_pivot_table ).

  " Generate Excel file
  CREATE OBJECT lo_writer.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
  
  MESSAGE 'Sales analysis pivot table created successfully' TYPE 'S'.
ENDMETHOD.
```

## Best Practices

### Design Guidelines

1. **Data Structure**: Ensure source data is properly normalized
2. **Field Selection**: Choose meaningful row, column, and value fields
3. **Aggregation**: Select appropriate aggregation functions for data types
4. **Performance**: Limit data range for large datasets

### User Experience

1. **Filtering**: Provide relevant filter options for interactivity
2. **Formatting**: Use consistent formatting and styling
3. **Labels**: Use clear, descriptive field names
4. **Layout**: Organize fields logically for easy interpretation

## Next Steps

After mastering pivot tables:

- **[Data Validation](/advanced/data-validation)** - Add input validation to source data
- **[Password Protection](/advanced/password-protection)** - Secure pivot table workbooks
- **[Macros](/advanced/macros)** - Automate pivot table operations

## Common Pivot Table Patterns

### Quick Reference

```abap
" Create pivot table
lo_pivot_table = lo_excel->add_new_table( ).
lo_pivot_table->settings-table_style = zcl_excel_table=>builtinstyle_pivot_light16.

" Configure field areas
ls_field-pivot_area = 'ROW'.     " Row field
ls_field-pivot_area = 'COLUMN'.  " Column field  
ls_field-pivot_area = 'DATA'.    " Value field
ls_field-pivot_area = 'PAGE'.    " Filter field

" Set aggregation functions
ls_field-totals_function = zcl_excel_table=>totals_function_sum.
ls_field-totals_function = zcl_excel_table=>totals_function_average.
ls_field-totals_function = zcl_excel_table=>totals_function_count.
```

This guide covers the comprehensive pivot table capabilities of abap2xlsx, enabling you to create sophisticated data analysis tools that provide interactive insights into your business data.
