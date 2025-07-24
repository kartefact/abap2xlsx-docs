# Templates

Advanced guide to creating and using Excel templates with abap2xlsx for consistent, reusable report structures.

## Understanding Excel Templates

Excel templates in abap2xlsx allow you to create predefined Excel structures that can be filled with dynamic data. This approach is particularly useful for standardized reports, forms, and documents that require consistent formatting and layout.

## Template Architecture

The template system in abap2xlsx is built around the `zcl_excel_fill_template` class <cite>src/zcl_excel_fill_template.clas.abap:31-90</cite>, which provides methods to load template files and populate them with data.

```abap
" Basic template usage
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_reader TYPE REF TO zif_excel_reader,
      lo_template_filler TYPE REF TO zcl_excel_fill_template,
      lv_template_file TYPE xstring.

" Load template file
CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
lo_excel = lo_reader->load_file( lv_template_file ).

" Create template filler
lo_template_filler = zcl_excel_fill_template=>create( lo_excel ).

" Fill template with data
lo_template_filler->fill_sheet( iv_data = ls_template_data ).
```

## Template Creation Strategies

### Design-First Approach

```abap
" Create template using Excel design tools first
METHOD create_design_first_template.
  " 1. Design template in Excel with placeholders
  " 2. Save as .xlsx file
  " 3. Load and populate in ABAP
  
  DATA: lo_reader TYPE REF TO zcl_excel_reader_2007,
        lo_excel TYPE REF TO zcl_excel,
        lv_template_path TYPE string VALUE '/templates/sales_report.xlsx'.

  CREATE OBJECT lo_reader.
  
  " Load pre-designed template
  lo_excel = lo_reader->load_file( 
    get_template_file_content( lv_template_path )
  ).

  " Template now ready for data population
  populate_template_data( lo_excel ).
ENDMETHOD.
```

### Code-First Approach

```abap
" Create template programmatically
METHOD create_code_first_template.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_template_style TYPE REF TO zcl_excel_style.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Report Template' ).

  " Create template structure
  create_template_header( lo_worksheet ).
  create_template_data_section( lo_worksheet ).
  create_template_footer( lo_worksheet ).

  " Save as template for reuse
  save_template( lo_excel 'sales_report_template.xlsx' ).
ENDMETHOD.
```

## Template Data Management

### Template Data Structure

The template system uses structured data types to organize information for template population:

```abap
" Define template data structure
TYPES: BEGIN OF ty_template_data,
         header TYPE BEGIN OF ty_header,
           title TYPE string,
           date TYPE d,
           company TYPE string,
           department TYPE string,
         END OF ty_header,
         
         data_table TYPE TABLE OF BEGIN OF ty_data_row,
           product TYPE string,
           quantity TYPE i,
           price TYPE p DECIMALS 2,
           total TYPE p DECIMALS 2,
         END OF ty_data_row,
         
         summary TYPE BEGIN OF ty_summary,
           total_items TYPE i,
           grand_total TYPE p DECIMALS 2,
           average_price TYPE p DECIMALS 2,
         END OF ty_summary,
       END OF ty_template_data.
```

### Data Population Methods

```abap
" Populate template with structured data
METHOD populate_template_with_data.
  DATA: ls_template_data TYPE ty_template_data,
        lo_template_filler TYPE REF TO zcl_excel_fill_template.

  " Prepare template data
  ls_template_data-header-title = 'Q4 Sales Report'.
  ls_template_data-header-date = sy-datum.
  ls_template_data-header-company = 'ACME Corporation'.
  ls_template_data-header-department = 'Sales'.

  " Add data rows
  APPEND VALUE #( product = 'Laptop' quantity = 10 price = '999.99' total = '9999.90' ) 
         TO ls_template_data-data_table.
  APPEND VALUE #( product = 'Mouse' quantity = 50 price = '29.99' total = '1499.50' ) 
         TO ls_template_data-data_table.

  " Calculate summary
  ls_template_data-summary-total_items = lines( ls_template_data-data_table ).
  ls_template_data-summary-grand_total = '11499.40'.
  ls_template_data-summary-average_price = '514.97'.

  " Fill template
  lo_template_filler = zcl_excel_fill_template=>create( lo_excel ).
  lo_template_filler->fill_sheet( ls_template_data ).
ENDMETHOD.
```

## Advanced Template Features

### Dynamic Template Sections

```abap
" Create templates with dynamic sections
METHOD create_dynamic_template_sections.
  DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lv_current_row TYPE i VALUE 1.

  " Header section (fixed)
  create_header_section( 
    io_worksheet = lo_worksheet
    iv_start_row = lv_current_row
  ).
  ADD 5 TO lv_current_row.

  " Dynamic data section (variable size)
  create_data_section(
    io_worksheet = lo_worksheet
    iv_start_row = lv_current_row
    it_data = lt_dynamic_data
    IMPORTING ev_end_row = lv_current_row
  ).
  ADD 2 TO lv_current_row.

  " Summary section (fixed)
  create_summary_section(
    io_worksheet = lo_worksheet
    iv_start_row = lv_current_row
  ).
ENDMETHOD.
```

### Template Inheritance

```abap
" Create template inheritance system
CLASS zcl_excel_template_base DEFINITION.
  PUBLIC SECTION.
    METHODS: create_base_template
               RETURNING VALUE(ro_excel) TYPE REF TO zcl_excel,
             apply_corporate_branding
               IMPORTING io_excel TYPE REF TO zcl_excel,
             add_standard_footer
               IMPORTING io_worksheet TYPE REF TO zcl_excel_worksheet.
ENDCLASS.

CLASS zcl_excel_sales_template DEFINITION INHERITING FROM zcl_excel_template_base.
  PUBLIC SECTION.
    METHODS: create_sales_template
               RETURNING VALUE(ro_excel) TYPE REF TO zcl_excel.
ENDCLASS.

CLASS zcl_excel_sales_template IMPLEMENTATION.
  METHOD create_sales_template.
    " Start with base template
    ro_excel = create_base_template( ).
    
    " Apply corporate branding
    apply_corporate_branding( ro_excel ).
    
    " Add sales-specific sections
    DATA(lo_worksheet) = ro_excel->get_active_worksheet( ).
    add_sales_specific_sections( lo_worksheet ).
    
    " Add standard footer
    add_standard_footer( lo_worksheet ).
  ENDMETHOD.
ENDCLASS.
```

## Template Placeholder System

### Placeholder Definition

```abap
" Define placeholder system for templates
METHOD define_template_placeholders.
  " Standard placeholder format: {{PLACEHOLDER_NAME}}
  CONSTANTS: BEGIN OF c_placeholders,
               company_name TYPE string VALUE '{{COMPANY_NAME}}',
               report_date TYPE string VALUE '{{REPORT_DATE}}',
               report_title TYPE string VALUE '{{REPORT_TITLE}}',
               user_name TYPE string VALUE '{{USER_NAME}}',
               data_table_start TYPE string VALUE '{{DATA_TABLE_START}}',
               data_table_end TYPE string VALUE '{{DATA_TABLE_END}}',
               total_amount TYPE string VALUE '{{TOTAL_AMOUNT}}',
             END OF c_placeholders.

  " Create placeholder mapping
  DATA: lt_placeholder_map TYPE TABLE OF BEGIN OF ty_placeholder,
          placeholder TYPE string,
          value TYPE string,
        END OF ty_placeholder.

  APPEND VALUE #( placeholder = c_placeholders-company_name 
                  value = 'ACME Corporation' ) TO lt_placeholder_map.
  APPEND VALUE #( placeholder = c_placeholders-report_date 
                  value = |{ sy-datum DATE = USER }| ) TO lt_placeholder_map.
  APPEND VALUE #( placeholder = c_placeholders-report_title 
                  value = 'Monthly Sales Report' ) TO lt_placeholder_map.

  " Apply placeholders to template
  replace_template_placeholders( 
    io_excel = lo_excel
    it_placeholder_map = lt_placeholder_map
  ).
ENDMETHOD.
```

### Placeholder Replacement Engine

```abap
" Replace placeholders in template
METHOD replace_template_placeholders.
  DATA: lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_worksheets TYPE REF TO zcl_excel_worksheets,
        lo_iterator TYPE REF TO zcl_excel_collection_iterator.

  " Iterate through all worksheets
  lo_worksheets = io_excel->get_worksheets( ).
  lo_iterator = lo_worksheets->get_iterator( ).

  WHILE lo_iterator->has_next( ) = abap_true.
    lo_worksheet = CAST zcl_excel_worksheet( lo_iterator->get_next( ) ).
    
    " Replace placeholders in worksheet
    replace_worksheet_placeholders(
      io_worksheet = lo_worksheet
      it_placeholder_map = it_placeholder_map
    ).
  ENDWHILE.
ENDMETHOD.

METHOD replace_worksheet_placeholders.
  DATA: lv_cell_value TYPE string,
        lv_max_row TYPE i,
        lv_max_col TYPE i.

  " Get worksheet dimensions
  lv_max_row = io_worksheet->get_highest_row( ).
  lv_max_col = io_worksheet->get_highest_column( ).

  " Scan all cells for placeholders
  DO lv_max_row TIMES.
    DATA(lv_row) = sy-index.
    
    DO lv_max_col TIMES.
      DATA(lv_col) = zcl_excel_common=>convert_column2alpha( sy-index ).
      
      lv_cell_value = io_worksheet->get_cell( 
        ip_column = lv_col 
        ip_row = lv_row 
      ).
      
      " Replace placeholders in cell value
      LOOP AT it_placeholder_map INTO DATA(ls_placeholder).
        REPLACE ALL OCCURRENCES OF ls_placeholder-placeholder 
                IN lv_cell_value 
                WITH ls_placeholder-value.
      ENDLOOP.
      
      " Update cell if changed
      IF lv_cell_value <> io_worksheet->get_cell( ip_column = lv_col ip_row = lv_row ).
        io_worksheet->set_cell(
          ip_column = lv_col
          ip_row = lv_row
          ip_value = lv_cell_value
        ).
      ENDIF.
    ENDDO.
  ENDDO.
ENDMETHOD.
```

## Template Library Management

### Template Repository

```abap
" Create template repository system
CLASS zcl_excel_template_repository DEFINITION.
  PUBLIC SECTION.
    TYPES: BEGIN OF ty_template_info,
             template_id TYPE string,
             name TYPE string,
             description TYPE string,
             category TYPE string,
             version TYPE string,
             created_by TYPE string,
             created_date TYPE d,
           END OF ty_template_info.

    METHODS: register_template
               IMPORTING is_template_info TYPE ty_template_info
                         io_template TYPE REF TO zcl_excel,
             get_template
               IMPORTING iv_template_id TYPE string
               RETURNING VALUE(ro_template) TYPE REF TO zcl_excel,
             list_templates
               IMPORTING iv_category TYPE string OPTIONAL
               RETURNING VALUE(rt_templates) TYPE TABLE OF ty_template_info.
ENDCLASS.
```

### Template Versioning

```abap
" Implement template versioning
METHOD manage_template_versions.
  DATA: ls_template_info TYPE zcl_excel_template_repository=>ty_template_info.

  " Version 1.0 - Basic sales template
  ls_template_info-template_id = 'SALES_REPORT_V1'.
  ls_template_info-name = 'Sales Report Template'.
  ls_template_info-version = '1.0'.
  ls_template_info-description = 'Basic sales report with summary'.
  
  " Version 2.0 - Enhanced with charts
  ls_template_info-template_id = 'SALES_REPORT_V2'.
  ls_template_info-version = '2.0'.
  ls_template_info-description = 'Enhanced sales report with charts and KPIs'.

  " Register templates
  go_template_repository->register_template( 
    is_template_info = ls_template_info
    io_template = create_sales_template_v2( )
  ).
ENDMETHOD.
```

## Template Generation Tools

### Code Generation from Templates

The library includes tools for generating ABAP code from Excel templates <cite>src/not_cloud/zexcel_template_get_types.prog.xml:1-58</cite>:

```abap
" Generate ABAP types from template structure
METHOD generate_types_from_template.
  " This functionality is provided by program ZEXCEL_TEMPLATE_GET_TYPES
  " which analyzes template structure and generates corresponding ABAP types
  
  " Example generated output:
  " TYPES: BEGIN OF ty_generated_template,
  "          header TYPE BEGIN OF ty_header,
  "            company_name TYPE string,
  "            report_date TYPE d,
  "          END OF ty_header,
  "          data_rows TYPE TABLE OF BEGIN OF ty_data_row,
  "            product TYPE string,
  "            amount TYPE p DECIMALS 2,
  "          END OF ty_data_row,
  "        END OF ty_generated_template.
ENDMETHOD.
```

### Template Validation

```abap
" Validate template structure and content
METHOD validate_template.
  DATA: lt_validation_errors TYPE TABLE OF string.

  " Check required placeholders
  validate_required_placeholders(
    io_excel = io_template
    IMPORTING et_errors = lt_validation_errors
  ).

  " Check template structure
  validate_template_structure(
    io_excel = io_template
    IMPORTING et_errors = lt_validation_errors
  ).

  " Check formatting consistency
  validate_formatting_consistency(
    io_excel = io_template
    IMPORTING et_errors = lt_validation_errors
  ).

  " Report validation results
  IF lt_validation_errors IS NOT INITIAL.
    LOOP AT lt_validation_errors INTO DATA(lv_error).
      MESSAGE lv_error TYPE 'W'.
    ENDLOOP.
  ELSE.
    MESSAGE 'Template validation successful' TYPE 'S'.
  ENDIF.
ENDMETHOD.
```

## Performance Optimization for Templates

### Template Caching

```abap
" Implement template caching for performance
CLASS zcl_excel_template_cache DEFINITION.
  PUBLIC SECTION.
    CLASS-DATA: gt_template_cache TYPE HASHED TABLE OF REF TO zcl_excel
                                   WITH UNIQUE KEY table_line.

    CLASS-METHODS: get_cached_template
                     IMPORTING iv_template_id TYPE string
                     RETURNING VALUE(ro_template) TYPE REF TO zcl_excel,
                   cache_template
                     IMPORTING iv_template_id TYPE string
                               io_template TYPE REF TO zcl_excel,
                   clear_cache.
ENDCLASS.

CLASS zcl_excel_template_cache IMPLEMENTATION.
  METHOD get_cached_template.
    " Check if template exists in cache
    READ TABLE gt_template_cache INTO ro_template 
         WITH KEY table_line = iv_template_id.
    
    IF ro_template IS NOT BOUND.
      " Load template from repository
      ro_template = load_template_from_repository( iv_template_id ).
      
      " Add to cache
      cache_template( 
        iv_template_id = iv_template_id
        io_template = ro_template
      ).
    ENDIF.
  ENDMETHOD.

  METHOD cache_template.
    INSERT io_template INTO TABLE gt_template_cache.
  ENDMETHOD.
ENDCLASS.
```

### Efficient Template Processing

```abap
" Optimize template processing for large datasets
METHOD process_template_efficiently.
  DATA: lo_template_filler TYPE REF TO zcl_excel_fill_template,
        lt_batch_data TYPE TABLE OF ty_template_data,
        lv_batch_size TYPE i VALUE 100.

  " Get cached template
  DATA(lo_template) = zcl_excel_template_cache=>get_cached_template( 'SALES_TEMPLATE' ).
  
  " Create template filler once
  lo_template_filler = zcl_excel_fill_template=>create( lo_template ).

  " Process data in batches
  DATA: lv_offset TYPE i VALUE 0.
  DO.
    " Get next batch
    SELECT * FROM data_table
      INTO TABLE lt_batch_data
      OFFSET lv_offset
      UP TO lv_batch_size ROWS.

    IF lt_batch_data IS INITIAL.
      EXIT.
    ENDIF.

    " Fill template with batch data
    LOOP AT lt_batch_data INTO DATA(ls_data).
      lo_template_filler->fill_sheet( ls_data ).
    ENDLOOP.

    " Clear batch and continue
    CLEAR lt_batch_data.
    ADD lv_batch_size TO lv_offset.
  ENDDO.
ENDMETHOD.
```

## Template Range Management

The template system uses range definitions to manage dynamic sections <cite>src/zcl_excel_fill_template.clas.abap:8-35</cite>:

```abap
" Working with template ranges
METHOD work_with_template_ranges.
  DATA: lo_template_filler TYPE REF TO zcl_excel_fill_template,
        lt_ranges TYPE zcl_excel_fill_template=>tt_ranges.

  " Create template filler
  lo_template_filler = zcl_excel_fill_template=>create( lo_excel ).
  
  " Access range information
  lt_ranges = lo_template_filler->mt_range.
  
  " Process each range
  LOOP AT lt_ranges INTO DATA(ls_range).
    MESSAGE |Range: { ls_range-name }, Sheet: { ls_range-sheet }, Start: { ls_range-start }, Stop: { ls_range-stop }| TYPE 'I'.
    
    " Handle different range types
    CASE ls_range-name.
      WHEN 'HEADER_RANGE'.
        process_header_range( ls_range ).
      WHEN 'DATA_RANGE'.
        process_data_range( ls_range ).
      WHEN 'SUMMARY_RANGE'.
        process_summary_range( ls_range ).
    ENDCASE.
  ENDLOOP.
ENDMETHOD.
```

## Complete Template Example

### Enterprise Report Template System

```abap
" Complete example: Enterprise template system
METHOD create_enterprise_template_system.
  DATA: lo_template_repository TYPE REF TO zcl_excel_template_repository,
        lo_template TYPE REF TO zcl_excel,
        lo_template_filler TYPE REF TO zcl_excel_fill_template,
        ls_template_data TYPE ty_enterprise_template_data.

  " Initialize template repository
  CREATE OBJECT lo_template_repository.

  " Get template from repository
  lo_template = lo_template_repository->get_template( 'ENTERPRISE_REPORT_V2' ).

  " Create template filler
  lo_template_filler = zcl_excel_fill_template=>create( lo_template ).

  " Prepare enterprise data
  prepare_enterprise_data( 
    IMPORTING es_data = ls_template_data 
  ).

  " Fill template with data
  TRY.
      lo_template_filler->fill_sheet( ls_template_data ).
      
      " Apply post-processing
      apply_template_post_processing( lo_template ).
      
      " Generate final file
      DATA: lo_writer TYPE REF TO zcl_excel_writer_2007.
      CREATE OBJECT lo_writer.
      DATA(lv_file) = lo_writer->write_file( lo_template ).
      
      " Handle file output
      process_generated_template_file( lv_file ).
      
    CATCH zcx_excel INTO DATA(lx_excel).
      MESSAGE |Template processing failed: { lx_excel->get_text( ) }| TYPE 'E'.
  ENDTRY.
ENDMETHOD.
```

## Best Practices for Templates

### Template Design Guidelines

1. **Consistency**: Use standardized placeholder naming conventions
2. **Flexibility**: Design templates to handle variable data sizes
3. **Maintainability**: Create modular template sections for easy updates
4. **Performance**: Cache frequently used templates

### Template Management Guidelines

1. **Version Control**: Implement proper versioning for template evolution
2. **Documentation**: Document template structure and data requirements
3. **Testing**: Validate templates with various data scenarios
4. **Repository**: Maintain centralized template repository

## Next Steps

After mastering templates:

- **[Automation](/advanced/automation)** - Automate template-based report generation
- **[Integration Patterns](/advanced/integration)** - Integrate templates with enterprise systems
- **[Performance Tuning](/advanced/performance-tuning)** - Optimize template processing
- **[Troubleshooting](/troubleshooting/template-issues)** - Diagnose template-related issues

## Common Template Patterns

### Quick Reference for Template Operations

```abap
" Load and fill template
lo_template_filler = zcl_excel_fill_template=>create( lo_excel ).
lo_template_filler->fill_sheet( ls_template_data ).

" Access template ranges
lt_ranges = lo_template_filler->mt_range.

" Generate types from template structure
" Use program ZEXCEL_TEMPLATE_GET_TYPES for code generation
```

This guide covers the comprehensive template capabilities of abap2xlsx. The template system provides powerful tools for creating consistent, reusable Excel reports that can be populated with dynamic data while maintaining professional formatting and structure.
