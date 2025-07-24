# ALV Integration

Comprehensive guide to converting ALV (ABAP List Viewer) grids to Excel format with abap2xlsx.

## Understanding ALV Integration

ALV integration in abap2xlsx allows you to convert existing ALV grids directly to Excel format, preserving formatting, filters, totals, and other ALV-specific features. The library provides several methods for ALV conversion depending on your system and requirements.

## Basic ALV Conversion

### Using bind_alv Method

The `bind_alv` method in the worksheet class provides direct ALV to Excel conversion <cite>src/zcl_excel_worksheet.clas.abap:138-147</cite>:

```abap
" Basic ALV to Excel conversion
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_alv_grid TYPE REF TO cl_gui_alv_grid,
      lt_data TYPE TABLE OF your_data_type.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).

" Convert ALV grid to Excel
lo_worksheet->bind_alv(
  io_alv = lo_alv_grid
  it_table = lt_data
  i_top = 1
  i_left = 1
  table_style = zcl_excel_table=>builtinstyle_medium9
  i_table = abap_true
).
```

### ALV with Table Styling

```abap
" Convert ALV with specific table styling
METHOD convert_alv_with_styling.
  DATA: lo_alv_grid TYPE REF TO cl_gui_alv_grid,
        lt_sales_data TYPE TABLE OF ty_sales_data.

  " Assume ALV grid is already populated with data
  
  lo_worksheet->bind_alv(
    io_alv = lo_alv_grid
    it_table = lt_sales_data
    i_top = 2                                    " Start from row 2
    i_left = 1                                   " Start from column A
    table_style = zcl_excel_table=>builtinstyle_light15
    i_table = abap_true                          " Create as Excel table
  ).
ENDMETHOD.
```

## Advanced ALV Integration

### Using ALV Converter Classes

The library includes specialized converter classes for ALV integration <cite>src/not_cloud/zcl_excel_converter.clas.xml:1-60</cite>:

```abap
" Using dedicated ALV converter
DATA: lo_alv_converter TYPE REF TO zcl_excel_converter_alv,
      lo_excel TYPE REF TO zcl_excel.

CREATE OBJECT lo_alv_converter.

" Configure ALV converter
lo_alv_converter->set_alv_grid( lo_alv_grid ).
lo_alv_converter->set_include_header( abap_true ).
lo_alv_converter->set_include_filters( abap_true ).
lo_alv_converter->set_include_totals( abap_true ).
lo_alv_converter->set_preserve_colors( abap_true ).

" Convert to Excel
lo_excel = lo_alv_converter->convert_to_excel( ).
```

### OLE2 Integration for Legacy Systems

For systems requiring OLE2 integration, the library provides specialized methods <cite>src/zcl_excel_worksheet.clas.abap:148-169</cite>:

```abap
" ALV to Excel using OLE2 (for legacy systems)
METHOD convert_alv_ole2.
  DATA: lo_alv_grid TYPE REF TO cl_gui_alv_grid,
        lt_listheader TYPE slis_t_listheader.

  " Configure list headers
  APPEND VALUE #( typ = 'H' info = 'Sales Report 2023' ) TO lt_listheader.
  APPEND VALUE #( typ = 'S' info = 'Generated on: ' && sy-datum ) TO lt_listheader.

  " Convert using OLE2 method
  lo_worksheet->bind_alv_ole2(
    i_save_path = 'C:\temp\sales_report.xlsx'
    io_alv = lo_alv_grid
    it_listheader = lt_listheader
    i_top = 1
    i_left = 1
    i_columns_header = 'X'
    i_columns_autofit = 'X'
  ).
ENDMETHOD.
```

## Preserving ALV Features

### Field Catalog Integration

```abap
" Preserve ALV field catalog information
METHOD preserve_alv_field_catalog.
  DATA: lt_fieldcat TYPE lvc_t_fcat,
        lt_excel_fieldcat TYPE zexcel_t_fieldcatalog,
        ls_excel_fieldcat TYPE zexcel_s_fieldcatalog.

  " Get ALV field catalog
  lo_alv_grid->get_frontend_fieldcatalog( 
    IMPORTING et_fieldcatalog = lt_fieldcat 
  ).

  " Convert to Excel field catalog
  LOOP AT lt_fieldcat INTO DATA(ls_fieldcat).
    CLEAR ls_excel_fieldcat.
    ls_excel_fieldcat-fieldname = ls_fieldcat-fieldname.
    ls_excel_fieldcat-position = ls_fieldcat-col_pos.
    ls_excel_fieldcat-datatype = ls_fieldcat-datatype.
    ls_excel_fieldcat-length = ls_fieldcat-outputlen.
    ls_excel_fieldcat-decimals = ls_fieldcat-decimals.
    
    " Handle conversion exits
    IF ls_fieldcat-convexit IS NOT INITIAL.
      ls_excel_fieldcat-convexit = ls_fieldcat-convexit.
    ENDIF.
    
    " Handle totals
    IF ls_fieldcat-do_sum = abap_true.
      ls_excel_fieldcat-totals_function = 'SUM'.
    ENDIF.
    
    APPEND ls_excel_fieldcat TO lt_excel_fieldcat.
  ENDLOOP.

  " Apply to worksheet
  lo_worksheet->bind_table(
    ip_table = lt_data
    it_field_catalog = lt_excel_fieldcat
  ).
ENDMETHOD.
```

### Preserving ALV Filters

```abap
" Convert ALV filters to Excel autofilters
METHOD preserve_alv_filters.
  DATA: lt_filter_index TYPE lvc_t_fidx,
        lo_autofilter TYPE REF TO zcl_excel_autofilter.

  " Get ALV filter information
  lo_alv_grid->get_filter_criteria(
    IMPORTING et_filter_index_table = lt_filter_index
  ).

  " Create Excel autofilter
  lo_autofilter = lo_worksheet->add_new_autofilter( ).
  lo_autofilter->set_range( 'A1:Z100' ).  " Adjust range as needed

  " Apply filter conditions
  LOOP AT lt_filter_index INTO DATA(ls_filter).
    " Convert ALV filter to Excel filter format
    convert_filter_condition(
      is_alv_filter = ls_filter
      io_autofilter = lo_autofilter
    ).
  ENDLOOP.
ENDMETHOD.
```

### Preserving ALV Totals and Subtotals

```abap
" Convert ALV totals to Excel formulas
METHOD preserve_alv_totals.
  DATA: lt_sort TYPE lvc_t_sort,
        lv_total_row TYPE i,
        lv_formula TYPE string.

  " Get ALV sort/total information
  lo_alv_grid->get_sort_criteria(
    IMPORTING et_sort = lt_sort
  ).

  " Find data range
  DATA(lv_last_row) = lo_worksheet->get_highest_row( ).
  lv_total_row = lv_last_row + 2.

  " Add total formulas based on ALV configuration
  LOOP AT lt_sort INTO DATA(ls_sort) WHERE spos > 0.
    " Add subtotal for each sort level
    IF ls_sort-subtot = abap_true.
      add_subtotal_formula(
        iv_column = ls_sort-fieldname
        iv_row = lv_total_row
        iv_function = 'SUM'
      ).
    ENDIF.
  ENDLOOP.

  " Add grand total
  lo_worksheet->set_cell(
    ip_column = 'A'
    ip_row = lv_total_row
    ip_value = 'Grand Total'
  ).
ENDMETHOD.
```

## Handling Different ALV Types

### Standard ALV Grid

```abap
" Convert standard ALV grid
METHOD convert_standard_alv.
  DATA: lo_alv_grid TYPE REF TO cl_gui_alv_grid.

  " Standard ALV conversion
  lo_worksheet->bind_alv(
    io_alv = lo_alv_grid
    it_table = lt_data
    i_top = 1
    i_left = 1
    table_style = zcl_excel_table=>builtinstyle_medium2
  ).
ENDMETHOD.
```

### Hierarchical ALV

```abap
" Convert hierarchical ALV to Excel
METHOD convert_hierarchical_alv.
  DATA: lo_alv_tree TYPE REF TO cl_gui_alv_tree,
        lt_hierarchy_data TYPE TABLE OF your_hierarchy_type.

  " For hierarchical ALV, manual conversion may be needed
  " Extract hierarchy information and convert to flat structure
  extract_hierarchy_data(
    io_alv_tree = lo_alv_tree
    IMPORTING et_flat_data = lt_hierarchy_data
  ).

  " Convert flattened data
  lo_worksheet->bind_table( ip_table = lt_hierarchy_data ).
ENDMETHOD.
```

### SALV Integration

```abap
" Convert SALV (Simple ALV) to Excel
METHOD convert_salv_to_excel.
  DATA: lo_salv_table TYPE REF TO cl_salv_table,
        lo_salv_converter TYPE REF TO zcl_excel_converter_salv_table.

  CREATE OBJECT lo_salv_converter.

  " Configure SALV converter
  lo_salv_converter->set_salv_table( lo_salv_table ).
  lo_salv_converter->set_include_aggregations( abap_true ).
  lo_salv_converter->set_include_layout( abap_true ).

  " Convert to Excel
  DATA(lo_excel) = lo_salv_converter->convert_to_excel( ).
ENDMETHOD.
```

## Performance Considerations

### Large ALV Datasets

```abap
" Handle large ALV datasets efficiently
METHOD convert_large_alv_dataset.
  DATA: lv_row_count TYPE i,
        lv_batch_size TYPE i VALUE 5000.

  " Check ALV data size
  lo_alv_grid->get_selected_rows(
    IMPORTING et_index_rows = DATA(lt_selected_rows)
  ).

  " If no selection, get total row count
  IF lt_selected_rows IS INITIAL.
    " Get total rows from ALV
    lv_row_count = get_alv_row_count( lo_alv_grid ).
  ELSE.
    lv_row_count = lines( lt_selected_rows ).
  ENDIF.

  " Use appropriate conversion method based on size
  IF lv_row_count > 10000.
    " Use huge file writer for very large datasets
    convert_alv_streaming( lo_alv_grid ).
  ELSE.
    " Use standard conversion
    lo_worksheet->bind_alv(
      io_alv = lo_alv_grid
      it_table = lt_data
    ).
  ENDIF.
ENDMETHOD.
```

### Memory Optimization

```abap
" Optimize memory usage during ALV conversion
METHOD optimize_alv_conversion.
  " Clear ALV selection to reduce memory usage
  lo_alv_grid->set_selected_rows( VALUE #( ) ).

  " Process in chunks if needed
  DATA: lv_start_row TYPE i VALUE 1,
        lv_end_row TYPE i,
        lv_chunk_size TYPE i VALUE 1000.

  DO.
    lv_end_row = lv_start_row + lv_chunk_size - 1.
    
    " Get data chunk from ALV
    DATA(lt_chunk) = get_alv_data_chunk(
      io_alv = lo_alv_grid
      iv_start = lv_start_row
      iv_end = lv_end_row
    ).

    IF lt_chunk IS INITIAL.
      EXIT.
    ENDIF.

    " Convert chunk
    convert_data_chunk( lt_chunk ).

    " Prepare for next chunk
    lv_start_row = lv_end_row + 1.
    CLEAR lt_chunk.
  ENDDO.
ENDMETHOD.
```

## Error Handling

### ALV Conversion Error Management

```abap
" Handle ALV conversion errors
METHOD handle_alv_conversion_errors.
  TRY.
      " Attempt ALV conversion
      lo_worksheet->bind_alv(
        io_alv = lo_alv_grid
        it_table = lt_data
        i_top = 1
        i_left = 1
      ).

    CATCH zcx_excel INTO DATA(lx_excel).
      " Handle Excel-specific errors
      MESSAGE |ALV conversion failed: { lx_excel->get_text( ) }| TYPE 'E'.

    CATCH cx_root INTO DATA(lx_root).
      " Handle general errors
      MESSAGE |Unexpected error during ALV conversion: { lx_root->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

## Complete ALV Integration Example

### Full ALV to Excel Report

```abap
" Complete example: Convert ALV report to Excel
METHOD create_alv_excel_report.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zif_excel_writer,
        lo_alv_grid TYPE REF TO cl_gui_alv_grid,
        lt_sales_data TYPE TABLE OF ty_sales_data.

  " Initialize Excel workbook
  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'ALV Export' ).

  " Add report header
  lo_worksheet->set_cell(
    ip_column = 'A'
    ip_row = 1
    ip_value = 'Sales Report - ALV Export'
  ).

  " Convert ALV with all features
  TRY.
      lo_worksheet->bind_alv(
        io_alv = lo_alv_grid
        it_table = lt_sales_data
        i_top = 3                                " Leave space for header
        i_left = 1
        table_style = zcl_excel_table=>builtinstyle_medium9
        i_table = abap_true
      ).

      " Preserve ALV-specific features
      preserve_alv_filters( ).
      preserve_alv_totals( ).
      preserve_alv_field_catalog( ).

      " Add metadata
      add_export_metadata( lo_worksheet ).

      " Generate Excel file
      CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
      DATA(lv_file) = lo_writer->write_file( lo_excel ).

      " Handle file output
      download_or_display_file( lv_file ).

    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |ALV export failed: { lx_excel->get_text( ) }| TYPE 'E'.
      
    CATCH cx_root INTO DATA(lx_root).
      MESSAGE |Unexpected error during ALV export: { lx_root->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

## Best Practices for ALV Integration

### Performance Guidelines

1. **Use Direct Binding**: The `bind_alv` method is optimized for ALV data structures
2. **Handle Large Datasets**: Use streaming approaches for very large ALV grids
3. **Memory Management**: Clear ALV selections and temporary data structures
4. **Batch Processing**: Process large ALV datasets in manageable chunks

### Feature Preservation Guidelines

1. **Field Catalogs**: Always preserve ALV field catalog information when possible
2. **Filters and Sorts**: Convert ALV filters to Excel autofilters for user convenience
3. **Totals and Subtotals**: Maintain calculation logic using Excel formulas
4. **Formatting**: Preserve ALV colors and formatting where supported

### Error Handling Guidelines

1. **Graceful Degradation**: Provide fallback options when full conversion isn't possible
2. **User Feedback**: Inform users about what features were preserved or lost
3. **Logging**: Maintain detailed logs of conversion processes for troubleshooting
4. **Validation**: Verify ALV grid state before attempting conversion

## Next Steps

After mastering ALV integration:

- **[Performance Optimization](/guide/performance)** - Optimize large ALV conversions
- **[Advanced Features](/advanced/custom-styles)** - Apply sophisticated formatting during ALV conversion
- **[Templates](/advanced/templates)** - Use templates for structured ALV reports
- **[Automation](/advanced/automation)** - Automate ALV to Excel conversion processes

## Common ALV Integration Patterns

### Quick Reference for ALV Operations

```abap
" Basic ALV conversion
lo_worksheet->bind_alv(
  io_alv = lo_alv_grid
  it_table = lt_data
  i_top = 1
  i_left = 1
  table_style = zcl_excel_table=>builtinstyle_medium9
  i_table = abap_true
).

" OLE2 conversion for legacy systems
lo_worksheet->bind_alv_ole2(
  i_save_path = 'C:\temp\report.xlsx'
  io_alv = lo_alv_grid
  it_listheader = lt_headers
  i_columns_header = 'X'
  i_columns_autofit = 'X'
).
```

This guide covers the comprehensive ALV integration capabilities of abap2xlsx. <cite>src/zcl_excel_worksheet.clas.abap:922-942</cite> The ALV integration system provides seamless conversion from ALV grids to Excel format while preserving as much functionality and formatting as possible.
