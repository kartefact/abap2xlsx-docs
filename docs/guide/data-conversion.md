I'll continue with the next file in the documentation structure. Here's the data conversion guide:

## `docs/guide/data-conversion.md`

```markdown
# Data Conversion

Comprehensive guide to converting ABAP data structures to Excel format with abap2xlsx.

## Understanding Data Conversion

Data conversion in abap2xlsx involves transforming ABAP internal tables, structures, and other data types into Excel-compatible formats. The library provides several approaches for this conversion, from simple table binding to sophisticated converter classes.

## Basic Table Binding

### Simple Internal Table to Excel

The most straightforward way to convert ABAP data to Excel is using the `bind_table` method [1](#29-0) :

```abap
" Define your data structure
TYPES: BEGIN OF ty_sales_data,
         region TYPE string,
         product TYPE string,
         quantity TYPE i,
         amount TYPE p DECIMALS 2,
         sale_date TYPE d,
       END OF ty_sales_data.

DATA: lt_sales_data TYPE TABLE OF ty_sales_data,
      lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

" Fill sample data
APPEND VALUE #( region = 'North' product = 'Laptop' quantity = 5 
                amount = '4999.95' sale_date = '20231201' ) TO lt_sales_data.
APPEND VALUE #( region = 'South' product = 'Mouse' quantity = 25 
                amount = '499.75' sale_date = '20231202' ) TO lt_sales_data.

" Create Excel and bind table
CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).

" Bind internal table to worksheet
lo_worksheet->bind_table(
  ip_table = lt_sales_data
  is_table_settings = VALUE #(
    top_left_column = 'A'
    top_left_row = 1
    table_style = zcl_excel_table=>builtinstyle_medium9
    show_row_stripes = abap_true
  )
).
```

### Advanced Table Settings

```abap
" Configure detailed table settings
DATA: ls_table_settings TYPE zexcel_s_table_settings.

ls_table_settings-top_left_column = 'B'.
ls_table_settings-top_left_row = 3.
ls_table_settings-table_style = zcl_excel_table=>builtinstyle_light15.
ls_table_settings-show_row_stripes = abap_true.
ls_table_settings-show_first_column = abap_true.
ls_table_settings-show_last_column = abap_false.
ls_table_settings-show_column_stripes = abap_false.

" Apply settings
lo_worksheet->bind_table(
  ip_table = lt_sales_data
  is_table_settings = ls_table_settings
).
```

## Data Type Conversion

### Automatic Type Detection

The worksheet handles automatic data type conversion based on ABAP field types [1](#29-0) :

```abap
" Different ABAP types are automatically converted
TYPES: BEGIN OF ty_mixed_data,
         text_field TYPE string,           " → Excel text
         number_field TYPE i,              " → Excel number
         decimal_field TYPE p DECIMALS 2,  " → Excel number with decimals
         date_field TYPE d,                " → Excel date
         time_field TYPE t,                " → Excel time
         boolean_field TYPE abap_bool,     " → Excel boolean
       END OF ty_mixed_data.

DATA: lt_mixed_data TYPE TABLE OF ty_mixed_data.

" The bind_table method automatically handles type conversion
lo_worksheet->bind_table( ip_table = lt_mixed_data ).
```

### Custom Field Conversion

```abap
" Handle special conversion requirements
METHOD convert_special_fields.
  DATA: lt_field_catalog TYPE zexcel_t_fieldcatalog,
        ls_field_catalog TYPE zexcel_s_fieldcatalog.

  " Define custom field conversion
  ls_field_catalog-fieldname = 'CUSTOMER_ID'.
  ls_field_catalog-convexit = 'ALPHA'.  " Apply ALPHA conversion exit
  APPEND ls_field_catalog TO lt_field_catalog.

  ls_field_catalog-fieldname = 'AMOUNT'.
  ls_field_catalog-convexit = cl_abap_typedescr=>typekind_float.
  APPEND ls_field_catalog TO lt_field_catalog.

  " Apply custom conversion during table binding
  lo_worksheet->bind_table(
    ip_table = lt_data
    it_field_catalog = lt_field_catalog
  ).
ENDMETHOD.
```

## Using Converter Classes

### Base Converter Class

The converter system provides sophisticated data transformation capabilities [2](#29-1) :

```abap
" Using the base converter class
DATA: lo_converter TYPE REF TO zcl_excel_converter,
      lo_excel TYPE REF TO zcl_excel.

CREATE OBJECT lo_converter.

" Configure converter settings
lo_converter->set_data_source( lt_internal_table ).
lo_converter->set_sheet_name( 'Converted Data' ).
lo_converter->set_include_header( abap_true ).

" Perform conversion
lo_excel = lo_converter->convert_to_excel( ).
```

### ALV Converter

```abap
" Convert ALV data to Excel
DATA: lo_alv_converter TYPE REF TO zcl_excel_converter_alv,
      lo_alv_grid TYPE REF TO cl_gui_alv_grid.

CREATE OBJECT lo_alv_converter.

" Set ALV grid as source
lo_alv_converter->set_alv_grid( lo_alv_grid ).

" Configure conversion options
lo_alv_converter->set_include_filters( abap_true ).
lo_alv_converter->set_include_totals( abap_true ).
lo_alv_converter->set_preserve_formatting( abap_true ).

" Convert to Excel
lo_excel = lo_alv_converter->convert_to_excel( ).
```

## Advanced Data Conversion Techniques

### Hierarchical Data Conversion

```abap
" Convert hierarchical data structures
METHOD convert_hierarchical_data.
  TYPES: BEGIN OF ty_header,
           order_id TYPE string,
           customer TYPE string,
           order_date TYPE d,
           items TYPE TABLE OF ty_item,
         END OF ty_header,
         BEGIN OF ty_item,
           item_id TYPE string,
           description TYPE string,
           quantity TYPE i,
           price TYPE p DECIMALS 2,
         END OF ty_item.

  DATA: lt_orders TYPE TABLE OF ty_header,
        lv_current_row TYPE i VALUE 1.

  " Process each order
  LOOP AT lt_orders INTO DATA(ls_order).
    " Add header information
    lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_current_row ip_value = ls_order-order_id ).
    lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_current_row ip_value = ls_order-customer ).
    lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_current_row ip_value = ls_order-order_date ).
    ADD 1 TO lv_current_row.

    " Add item details
    LOOP AT ls_order-items INTO DATA(ls_item).
      lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_current_row ip_value = ls_item-item_id ).
      lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_current_row ip_value = ls_item-description ).
      lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_current_row ip_value = ls_item-quantity ).
      lo_worksheet->set_cell( ip_column = 'E' ip_row = lv_current_row ip_value = ls_item-price ).
      ADD 1 TO lv_current_row.
    ENDLOOP.

    " Add separator row
    ADD 1 TO lv_current_row.
  ENDLOOP.
ENDMETHOD.
```

### Dynamic Field Mapping

```abap
" Dynamic field mapping based on runtime information
METHOD convert_with_dynamic_mapping.
  DATA: lo_struct_descr TYPE REF TO cl_abap_structdescr,
        lt_components TYPE cl_abap_structdescr=>component_table,
        lv_column TYPE string VALUE 'A',
        lv_row TYPE i VALUE 1.

  " Get structure description
  lo_struct_descr ?= cl_abap_typedescr=>describe_by_data( <ls_data> ).
  lt_components = lo_struct_descr->get_components( ).

  " Create headers based on field names
  LOOP AT lt_components INTO DATA(ls_component).
    lo_worksheet->set_cell( 
      ip_column = lv_column 
      ip_row = 1 
      ip_value = ls_component-name 
    ).
    lv_column = zcl_excel_common=>convert_column2alpha( 
      zcl_excel_common=>convert_column2int( lv_column ) + 1 
    ).
  ENDLOOP.

  " Add data rows
  lv_row = 2.
  LOOP AT lt_data ASSIGNING <ls_data>.
    lv_column = 'A'.
    LOOP AT lt_components INTO ls_component.
      ASSIGN COMPONENT ls_component-name OF STRUCTURE <ls_data> TO <lv_field>.
      IF sy-subrc = 0.
        lo_worksheet->set_cell( 
          ip_column = lv_column 
          ip_row = lv_row 
          ip_value = <lv_field> 
        ).
      ENDIF.
      lv_column = zcl_excel_common=>convert_column2alpha( 
        zcl_excel_common=>convert_column2int( lv_column ) + 1 
      ).
    ENDLOOP.
    ADD 1 TO lv_row.
  ENDLOOP.
ENDMETHOD.
```

## Data Validation and Cleansing

### Input Data Validation

```abap
" Validate data before conversion
METHOD validate_conversion_data.
  DATA: lv_errors TYPE i.

  " Check for required fields
  LOOP AT lt_data INTO DATA(ls_data).
    IF ls_data-key_field IS INITIAL.
      MESSAGE |Row { sy-tabix }: Key field is empty| TYPE 'E'.
      ADD 1 TO lv_errors.
    ENDIF.

    " Validate data ranges
    IF ls_data-amount < 0.
      MESSAGE |Row { sy-tabix }: Negative amount not allowed| TYPE 'W'.
    ENDIF.

    " Check date validity
    IF ls_data-date_field IS NOT INITIAL.
      CALL FUNCTION 'DATE_CHECK_PLAUSIBILITY'
        EXPORTING
          date = ls_data-date_field
        EXCEPTIONS
          plausibility_check_failed = 1.
      IF sy-subrc <> 0.
        MESSAGE |Row { sy-tabix }: Invalid date| TYPE 'E'.
        ADD 1 TO lv_errors.
      ENDIF.
    ENDIF.
  ENDLOOP.

  IF lv_errors > 0.
    MESSAGE |{ lv_errors } validation errors found| TYPE 'E'.
  ENDIF.
ENDMETHOD.
```

### Data Cleansing

```abap
" Clean and normalize data before conversion
METHOD cleanse_data_for_excel.
  LOOP AT lt_data ASSIGNING <ls_data>.
    " Remove leading zeros from text fields
    IF <ls_data>-text_field CA '0123456789'.
      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
        EXPORTING
          input = <ls_data>-text_field
        IMPORTING
          output = <ls_data>-text_field.
    ENDIF.

    " Normalize decimal separators
    REPLACE ALL OCCURRENCES OF ',' IN <ls_data>-amount_text WITH '.'.

    " Trim whitespace
    CONDENSE <ls_data>-description.

    " Handle special characters for Excel compatibility
    REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf 
            IN <ls_data>-notes WITH ' '.
  ENDLOOP.
ENDMETHOD.
```

## Performance Optimization

### Batch Processing

```abap
" Process large datasets in batches
METHOD convert_large_dataset.
  CONSTANTS: c_batch_size TYPE i VALUE 1000.
  
  DATA: lv_offset TYPE i,
        lv_current_row TYPE i VALUE 2,
        lt_batch TYPE TABLE OF ty_data.

  " Process data in batches
  DO.
    " Get next batch
    SELECT * FROM source_table
      INTO TABLE lt_batch
      OFFSET lv_offset
      UP TO c_batch_size ROWS.

    IF lt_batch IS INITIAL.
      EXIT.
    ENDIF.

    " Convert batch to Excel
    LOOP AT lt_batch INTO DATA(ls_data).
      " Add row to worksheet
      add_data_row( 
        is_data = ls_data 
        iv_row = lv_current_row 
      ).
      ADD 1 TO lv_current_row.
    ENDLOOP.

    " Prepare for next batch
    ADD c_batch_size TO lv_offset.
    CLEAR lt_batch.

    " Optional: Commit work for long-running processes
    COMMIT WORK.
  ENDDO.
ENDMETHOD.
```

### Memory-Efficient Conversion

```abap
" Use streaming approach for very large datasets
METHOD stream_data_to_excel.
  DATA: lo_huge_writer TYPE REF TO zcl_excel_writer_huge_file,
        lv_row TYPE i VALUE 1.

  " Use huge file writer for memory efficiency
  CREATE OBJECT lo_huge_writer.

  " Add headers
  add_header_row( ).

  " Stream data row by row
  SELECT * FROM large_table INTO DATA(ls_data).
    ADD 1 TO lv_row.
    
    " Add data directly to writer
    lo_huge_writer->add_row(
      ip_row = lv_row
      ip_data = ls_data
    ).

    " Periodic memory cleanup
    IF lv_row MOD 10000 = 0.
      lo_huge_writer->flush_buffer( ).
    ENDIF.
  ENDSELECT.

  " Generate final file
  DATA(lv_file) = lo_huge_writer->write_file( lo_excel ).
ENDMETHOD.
```

## Error Handling and Recovery

### Conversion Error Management

```abap
" Handle conversion errors gracefully
METHOD convert_with_error_handling.
  DATA: lt_error_log TYPE TABLE OF string,
        lv_success_count TYPE i,
        lv_error_count TYPE i.

  LOOP AT lt_source_data INTO DATA(ls_source).
    TRY.
        " Attempt conversion
        convert_single_record( 
          is_source = ls_source
          iv_target_row = sy-tabix + 1
        ).
        ADD 1 TO lv_success_count.

      CATCH zcx_excel INTO DATA(lx_excel).
        ADD 1 TO lv_error_count.
        APPEND |Row { sy-tabix }: { lx_excel->get_text( ) }| TO lt_error_log.
        
        " Continue with next record or abort based on error severity
        IF lv_error_count > 10.
          MESSAGE 'Too many conversion errors - aborting' TYPE 'E'.
        ENDIF.
        
      CATCH cx_root INTO DATA(lx_root).
        ADD 1 TO lv_error_count.
        APPEND |Row { sy-tabix }: Unexpected error - { lx_root->get_text( ) }| TO lt_error_log.
    ENDTRY.
  ENDLOOP.

  " Report conversion results
  MESSAGE |Conversion complete: { lv_success_count } successful, { lv_error_count } errors| TYPE 'I'.
  
  " Log errors for review
  IF lt_error_log IS NOT INITIAL.
    write_error_log( lt_error_log ).
  ENDIF.
ENDMETHOD.
```

### Data Recovery Strategies

```abap
" Implement fallback strategies for problematic data
METHOD convert_with_fallbacks.
  LOOP AT lt_data ASSIGNING <ls_data>.
    " Primary conversion attempt
    TRY.
        convert_field_primary( <ls_data> ).
        
      CATCH zcx_excel.
        " Fallback 1: Use default values
        TRY.
            convert_field_with_defaults( <ls_data> ).
            
          CATCH zcx_excel.
            " Fallback 2: Skip problematic fields
            convert_field_minimal( <ls_data> ).
        ENDTRY.
    ENDTRY.
  ENDLOOP.
ENDMETHOD.
```

## Conversion Exit Handling

The worksheet conversion system includes sophisticated handling of ABAP conversion exits <cite>src/zcl_excel_worksheet.clas.abap:2334-2347</cite>:

```abap
" Conversion exits are automatically applied during data conversion
METHOD handle_conversion_exits.
  " ALPHA conversion exit for leading zeros
  IF ls_field_catalog-convexit = 'ALPHA'.
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
      EXPORTING
        input = lv_raw_value
      IMPORTING
        output = lv_formatted_value.
  ENDIF.
  
  " Numeric conversion for proper Excel formatting
  IF ls_field_catalog-convexit = cl_abap_typedescr=>typekind_float.
    lv_number = zcl_excel_common=>excel_string_to_number( lv_raw_value ).
    lv_formatted_value = |{ lv_number NUMBER = RAW }|.
  ENDIF.
ENDMETHOD.
```

## Complete Data Conversion Example

### Enterprise Data Export System

```abap
" Complete example: Enterprise data conversion with error handling
METHOD create_enterprise_export.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zif_excel_writer,
        lt_sales_data TYPE TABLE OF ty_sales_record,
        lt_field_catalog TYPE zexcel_t_fieldcatalog.

  " Initialize Excel workbook
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Export' ).

  " Load and validate data
  load_sales_data( 
    IMPORTING et_data = lt_sales_data 
  ).
  
  validate_conversion_data( lt_sales_data ).
  cleanse_data_for_excel( CHANGING ct_data = lt_sales_data ).

  " Configure field catalog for special conversions
  prepare_field_catalog( 
    IMPORTING et_catalog = lt_field_catalog 
  ).

  " Perform conversion with error handling
  TRY.
      lo_worksheet->bind_table(
        ip_table = lt_sales_data
        it_field_catalog = lt_field_catalog
        is_table_settings = VALUE #(
          top_left_column = 'A'
          top_left_row = 2
          table_style = zcl_excel_table=>builtinstyle_medium9
          show_row_stripes = abap_true
          show_first_column = abap_true
        )
      ).

      " Add header information
      add_report_header( lo_worksheet ).
      
      " Add summary calculations
      add_summary_formulas( lo_worksheet ).

      " Generate file
      CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
      DATA(lv_file) = lo_writer->write_file( lo_excel ).
      
      " Handle file output (download, save, email, etc.)
      process_generated_file( lv_file ).
      
    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |Export failed: { lx_excel->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

## Best Practices for Data Conversion

### Performance Guidelines

1. **Use Table Binding**: Prefer `bind_table` over individual cell operations for large datasets
2. **Batch Processing**: Process large datasets in manageable chunks
3. **Memory Management**: Clear objects and variables when no longer needed
4. **Efficient Writers**: Use `zcl_excel_writer_huge_file` for very large files

### Data Quality Guidelines

1. **Validate Input**: Always validate data before conversion
2. **Handle Nulls**: Provide appropriate handling for empty/null values
3. **Type Consistency**: Ensure consistent data types within columns
4. **Error Logging**: Implement comprehensive error logging and reporting

### Maintainability Guidelines

1. **Modular Design**: Break conversion logic into reusable methods
2. **Configuration**: Use field catalogs and settings for flexibility
3. **Documentation**: Document conversion rules and business logic
4. **Testing**: Implement unit tests for conversion methods

## Next Steps

After mastering data conversion:

- **[ALV Integration](/guide/alv-integration)** - Convert ALV grids to Excel format
- **[Performance Optimization](/guide/performance)** - Optimize large data conversions
- **[Advanced Features](/advanced/custom-styles)** - Apply sophisticated formatting during conversion
- **[Templates](/advanced/templates)** - Use templates for structured data presentation

## Common Conversion Patterns

### Quick Reference for Data Operations

```abap
" Basic table binding
lo_worksheet->bind_table( ip_table = lt_data ).

" With field catalog
lo_worksheet->bind_table( 
  ip_table = lt_data 
  it_field_catalog = lt_field_catalog 
).

" With table settings
lo_worksheet->bind_table(
  ip_table = lt_data
  is_table_settings = ls_settings
).
```

This guide covers the comprehensive data conversion capabilities of abap2xlsx. The conversion system handles automatic type detection, field formatting, and provides extensive customization options for transforming ABAP data into professional Excel reports.
