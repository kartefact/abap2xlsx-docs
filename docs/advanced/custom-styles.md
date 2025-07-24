# Custom Styles

Advanced guide to creating sophisticated custom styles and style templates with abap2xlsx.

## Advanced Style Architecture

Building on the basic styling concepts, advanced custom styles allow you to create reusable style templates, dynamic styling systems, and complex formatting rules that can be applied consistently across your Excel reports.

```abap
" Advanced style template system
CLASS zcl_excel_style_factory DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: create_corporate_header
                     IMPORTING io_excel TYPE REF TO zcl_excel
                     RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style,
                   create_data_row_alternating
                     IMPORTING io_excel TYPE REF TO zcl_excel
                               iv_row_number TYPE i
                     RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style,
                   create_conditional_style
                     IMPORTING io_excel TYPE REF TO zcl_excel
                               iv_condition TYPE string
                               iv_value TYPE any
                     RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style.
ENDCLASS.
```

## Style Template System

### Corporate Style Templates

<cite>src/zcl_excel.clas.abap:481-540</cite>

```abap
" Create comprehensive corporate style system
METHOD create_corporate_style_system.
  DATA: lo_title_style TYPE REF TO zcl_excel_style,
        lo_header_style TYPE REF TO zcl_excel_style,
        lo_subheader_style TYPE REF TO zcl_excel_style,
        lo_data_style TYPE REF TO zcl_excel_style,
        lo_total_style TYPE REF TO zcl_excel_style,
        lo_highlight_style TYPE REF TO zcl_excel_style.

  " Corporate title style
  lo_title_style = lo_excel->add_new_style( ).
  lo_title_style->font->name = 'Arial'.
  lo_title_style->font->size = 18.
  lo_title_style->font->bold = abap_true.
  lo_title_style->font->color->set_rgb( '1F4E79' ).  " Corporate blue
  lo_title_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  lo_title_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.

  " Corporate header style
  lo_header_style = lo_excel->add_new_style( ).
  lo_header_style->font->name = 'Arial'.
  lo_header_style->font->size = 12.
  lo_header_style->font->bold = abap_true.
  lo_header_style->font->color->set_rgb( 'FFFFFF' ).
  lo_header_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_header_style->fill->fgcolor->set_rgb( '4472C4' ).
  lo_header_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
  lo_header_style->borders->allborders->border_color->set_rgb( '2F5597' ).
  lo_header_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.

  " Store styles in factory for reuse
  register_corporate_style( 'TITLE' lo_title_style ).
  register_corporate_style( 'HEADER' lo_header_style ).
ENDMETHOD.
```

### Dynamic Style Generation

```abap
" Generate styles based on data conditions
METHOD create_dynamic_styles.
  DATA: lo_style TYPE REF TO zcl_excel_style,
        lv_color TYPE string.

  " Create styles based on performance thresholds
  LOOP AT lt_performance_data INTO DATA(ls_data).
    " Determine color based on performance
    CASE ls_data-performance_rating.
      WHEN 'EXCELLENT'.
        lv_color = '00B050'.  " Green
      WHEN 'GOOD'.
        lv_color = 'FFC000'.  " Orange
      WHEN 'POOR'.
        lv_color = 'FF0000'.  " Red
      WHEN OTHERS.
        lv_color = 'D9D9D9'.  " Gray
    ENDCASE.

    " Create or retrieve cached style
    lo_style = get_or_create_performance_style( 
      iv_color = lv_color
      iv_rating = ls_data-performance_rating
    ).

    " Apply to cell
    lo_worksheet->set_cell(
      ip_column = 'E'
      ip_row = sy-tabix + 1
      ip_value = ls_data-performance_rating
      ip_style = lo_style
    ).
  ENDLOOP.
ENDMETHOD.
```

## Advanced Conditional Formatting

### Complex Conditional Rules

```abap
" Create sophisticated conditional formatting rules
METHOD create_advanced_conditional_formatting.
  DATA: lo_cond_format TYPE REF TO zcl_excel_style_cond,
        lo_style TYPE REF TO zcl_excel_style.

  " Multi-condition formatting for financial data
  lo_cond_format = lo_worksheet->add_new_style_cond( ).
  lo_cond_format->set_range( 'D2:D1000' ).
  
  " Rule 1: Highlight negative values in red
  lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_cond_format->set_operator( zcl_excel_style_cond=>c_operator_lessthan ).
  lo_cond_format->set_formula( '0' ).
  lo_cond_format->set_color( zcl_excel_style_color=>c_red ).
  lo_cond_format->set_font_bold( abap_true ).

  " Rule 2: Data bars for positive values
  DATA(lo_databar_format) = lo_worksheet->add_new_style_cond( ).
  lo_databar_format->set_range( 'D2:D1000' ).
  lo_databar_format->set_rule_type( zcl_excel_style_cond=>c_rule_databar ).
  lo_databar_format->set_color( zcl_excel_style_color=>c_blue ).
  lo_databar_format->set_formula( 'AND(D2>0,D2<MAX($D$2:$D$1000))' ).

  " Rule 3: Icon sets for trend analysis
  DATA(lo_iconset_format) = lo_worksheet->add_new_style_cond( ).
  lo_iconset_format->set_range( 'E2:E1000' ).
  lo_iconset_format->set_rule_type( zcl_excel_style_cond=>c_rule_iconset ).
  lo_iconset_format->set_iconset_type( zcl_excel_style_cond=>c_iconset_3arrows ).
ENDMETHOD.
```

### Formula-Based Conditional Formatting

```abap
" Use Excel formulas for complex conditional logic
METHOD create_formula_based_conditions.
  DATA: lo_cond_format TYPE REF TO zcl_excel_style_cond.

  " Highlight entire rows based on status
  lo_cond_format = lo_worksheet->add_new_style_cond( ).
  lo_cond_format->set_range( 'A2:Z1000' ).
  lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_cond_format->set_formula( '$F2="OVERDUE"' ).
  lo_cond_format->set_fill_color( 'FFE6E6' ).  " Light red background

  " Highlight duplicates across multiple columns
  DATA(lo_duplicate_format) = lo_worksheet->add_new_style_cond( ).
  lo_duplicate_format->set_range( 'A2:C1000' ).
  lo_duplicate_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_duplicate_format->set_formula( 'COUNTIFS($A$2:$A$1000,$A2,$B$2:$B$1000,$B2,$C$2:$C$1000,$C2)>1' ).
  lo_duplicate_format->set_fill_color( 'FFFF99' ).  " Yellow background

  " Top/Bottom percentage formatting
  DATA(lo_percentile_format) = lo_worksheet->add_new_style_cond( ).
  lo_percentile_format->set_range( 'D2:D1000' ).
  lo_percentile_format->set_rule_type( zcl_excel_style_cond=>c_rule_top10 ).
  lo_percentile_format->set_rank( 10 ).
  lo_percentile_format->set_percent( abap_true ).
  lo_percentile_format->set_fill_color( '00FF00' ).  " Green for top 10%
ENDMETHOD.
```

## Style Inheritance and Cascading

### Hierarchical Style System

```abap
" Create hierarchical style inheritance
CLASS zcl_excel_style_hierarchy DEFINITION.
  PUBLIC SECTION.
    TYPES: BEGIN OF ty_style_config,
             base_style TYPE string,
             font_name TYPE string,
             font_size TYPE i,
             font_color TYPE string,
             fill_color TYPE string,
             border_style TYPE string,
           END OF ty_style_config.

    METHODS: create_inherited_style
               IMPORTING is_config TYPE ty_style_config
                         io_parent_style TYPE REF TO zcl_excel_style OPTIONAL
               RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style.
ENDCLASS.

CLASS zcl_excel_style_hierarchy IMPLEMENTATION.
  METHOD create_inherited_style.
    " Create new style inheriting from parent
    ro_style = lo_excel->add_new_style( ).
    
    " Inherit from parent if provided
    IF io_parent_style IS BOUND.
      copy_style_properties( 
        io_source = io_parent_style
        io_target = ro_style
      ).
    ENDIF.

    " Apply specific overrides
    IF is_config-font_name IS NOT INITIAL.
      ro_style->font->name = is_config-font_name.
    ENDIF.
    
    IF is_config-font_size > 0.
      ro_style->font->size = is_config-font_size.
    ENDIF.
    
    IF is_config-font_color IS NOT INITIAL.
      ro_style->font->color->set_rgb( is_config-font_color ).
    ENDIF.
  ENDMETHOD.
ENDCLASS.
```

### Style Composition

```abap
" Compose complex styles from simpler components
METHOD compose_complex_styles.
  DATA: lo_base_style TYPE REF TO zcl_excel_style,
        lo_font_component TYPE REF TO zcl_excel_style_font,
        lo_fill_component TYPE REF TO zcl_excel_style_fill,
        lo_border_component TYPE REF TO zcl_excel_style_borders.

  " Create base style
  lo_base_style = lo_excel->add_new_style( ).

  " Compose font component
  CREATE OBJECT lo_font_component.
  lo_font_component->name = 'Calibri'.
  lo_font_component->size = 11.
  lo_font_component->color->set_theme( zcl_excel_style_color=>c_theme_dark1 ).

  " Compose fill component
  CREATE OBJECT lo_fill_component.
  lo_fill_component->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_fill_component->fgcolor->set_rgb( 'F2F2F2' ).

  " Compose border component
  CREATE OBJECT lo_border_component.
  lo_border_component->allborders->border_style = zcl_excel_style_border=>c_border_thin.
  lo_border_component->allborders->border_color->set_rgb( 'BFBFBF' ).

  " Assign components to style
  lo_base_style->font = lo_font_component.
  lo_base_style->fill = lo_fill_component.
  lo_base_style->borders = lo_border_component.
ENDMETHOD.
```

## Performance-Optimized Styling

### Style Caching System

```abap
" Implement efficient style caching
CLASS zcl_excel_style_cache DEFINITION.
  PUBLIC SECTION.
    TYPES: BEGIN OF ty_style_key,
             font_name TYPE string,
             font_size TYPE i,
             font_bold TYPE abap_bool,
             fill_color TYPE string,
             border_style TYPE string,
           END OF ty_style_key.

    TYPES: tt_style_cache TYPE HASHED TABLE OF REF TO zcl_excel_style
                            WITH UNIQUE KEY table_line.

    CLASS-DATA: gt_style_cache TYPE tt_style_cache.

    CLASS-METHODS: get_cached_style
                     IMPORTING is_key TYPE ty_style_key
                               io_excel TYPE REF TO zcl_excel
                     RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style,
                   clear_cache.
ENDCLASS.

CLASS zcl_excel_style_cache IMPLEMENTATION.
  METHOD get_cached_style.
    " Generate cache key
    DATA(lv_cache_key) = generate_cache_key( is_key ).
    
    " Check if style exists in cache
    READ TABLE gt_style_cache INTO ro_style WITH KEY table_line = lv_cache_key.
    
    IF ro_style IS NOT BOUND.
      " Create new style
      ro_style = create_style_from_key( is_key io_excel ).
      
      " Add to cache
      INSERT ro_style INTO TABLE gt_style_cache.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
```

### Bulk Style Application

```abap
" Apply styles efficiently to large ranges
METHOD apply_bulk_styles.
  DATA: lt_style_ranges TYPE TABLE OF zexcel_s_style_range,
        ls_style_range TYPE zexcel_s_style_range.

  " Define style ranges for bulk application
  ls_style_range-range = 'A1:Z1'.
  ls_style_range-style = get_cached_style( is_header_key ).
  APPEND ls_style_range TO lt_style_ranges.

  ls_style_range-range = 'A2:Z1000'.
  ls_style_range-style = get_cached_style( is_data_key ).
  APPEND ls_style_range TO lt_style_ranges.

  ls_style_range-range = 'A1001:Z1001'.
  ls_style_range-style = get_cached_style( is_total_key ).
  APPEND ls_style_range TO lt_style_ranges.

  " Apply all styles in one operation
  lo_worksheet->set_styles_bulk( lt_style_ranges ).
ENDMETHOD.
```

## Theme Integration

### Custom Theme Creation

```abap
" Create custom Excel themes
METHOD create_custom_theme.
  DATA: lo_theme TYPE REF TO zcl_excel_theme.

  " Create new theme instance
  CREATE OBJECT lo_theme.
  lo_theme->set_theme_name( 'Corporate Theme 2024' ).

  " Configure color scheme
  lo_theme->set_color_scheme_name( 'Corporate Colors' ).
  lo_theme->set_color( 
    iv_type = 'accent1' 
    iv_srgb = '4472C4'  " Corporate blue
  ).
  lo_theme->set_color( 
    iv_type = 'accent2' 
    iv_srgb = '70AD47'  " Corporate green
  ).
  lo_theme->set_color( 
    iv_type = 'accent3' 
    iv_srgb = 'FFC000'  " Corporate orange
  ).

  " Configure font scheme
  lo_theme->set_font_scheme_name( 'Corporate Fonts' ).
  lo_theme->set_latin_font( 
    iv_type = 'majorFont'
    iv_typeface = 'Arial'
  ).
  lo_theme->set_latin_font( 
    iv_type = 'minorFont'
    iv_typeface = 'Calibri'
  ).

  " Apply theme to workbook
  lo_excel->set_theme( lo_theme ).
ENDMETHOD.
```

<cite>src/zcl_excel_theme.clas.abap:18-66</cite>

### Theme-Based Style Creation

```abap
" Create styles that reference theme colors
METHOD create_theme_based_styles.
  DATA: lo_style TYPE REF TO zcl_excel_style.

  " Header style using theme colors
  lo_style = lo_excel->add_new_style( ).
  lo_style->font->color->set_theme( zcl_excel_style_color=>c_theme_light1 ).
  lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_style->fill->fgcolor->set_theme( zcl_excel_style_color=>c_theme_accent1 ).

  " Data style with theme-based alternating colors
  DATA(lo_alt_style) = lo_excel->add_new_style( ).
  lo_alt_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_alt_style->fill->fgcolor->set_theme( zcl_excel_style_color=>c_theme_accent2 ).
  lo_alt_style->fill->fgcolor->set_tint( '0.8' ).  " Lighter shade
ENDMETHOD.
```

<cite>src/zcl_excel_theme_color_scheme.clas.abap:88-300</cite>

## Advanced Style Patterns

### Responsive Style System

```abap
" Create styles that adapt to content
METHOD create_responsive_styles.
  DATA: lo_style TYPE REF TO zcl_excel_style,
        lv_content_length TYPE i.

  " Adjust font size based on content length
  LOOP AT lt_data INTO DATA(ls_data).
    lv_content_length = strlen( ls_data-description ).
    
    " Create style based on content characteristics
    lo_style = lo_excel->add_new_style( ).
    
    CASE lv_content_length.
      WHEN 0 TO 20.
        lo_style->font->size = 11.
      WHEN 21 TO 50.
        lo_style->font->size = 10.
      WHEN OTHERS.
        lo_style->font->size = 9.
        lo_style->alignment->wraptext = abap_true.
    ENDCASE.

    " Apply to cell
    lo_worksheet->set_cell(
      ip_column = 'C'
      ip_row = sy-tabix + 1
      ip_value = ls_data-description
      ip_style = lo_style
    ).
  ENDLOOP.
ENDMETHOD.
```

### Style Animation and Effects

```abap
" Create styles with visual effects
METHOD create_effect_styles.
  DATA: lo_gradient_style TYPE REF TO zcl_excel_style,
        lo_shadow_style TYPE REF TO zcl_excel_style.

  " Gradient fill style
  lo_gradient_style = lo_excel->add_new_style( ).
  lo_gradient_style->fill->filltype = zcl_excel_style_fill=>c_fill_gradient_linear.
  lo_gradient_style->fill->fgcolor->set_rgb( '4472C4' ).
  lo_gradient_style->fill->bgcolor->set_rgb( 'FFFFFF' ).
  lo_gradient_style->fill->gradient_degree = 90.  " Vertical gradient

  " Shadow effect style (simulated with borders)
  lo_shadow_style = lo_excel->add_new_style( ).
  lo_shadow_style->borders->right->border_style = zcl_excel_style_border=>c_border_medium.
  lo_shadow_style->borders->right->border_color->set_rgb( 'CCCCCC' ).
  lo_shadow_style->borders->bottom->border_style = zcl_excel_style_border=>c_border_medium.
  lo_shadow_style->borders->bottom->border_color->set_rgb( 'CCCCCC' ).
ENDMETHOD.
```

## Style Validation and Quality Control

### Style Consistency Checker

```abap
" Validate style consistency across workbook
METHOD validate_style_consistency.
  DATA: lt_used_styles TYPE TABLE OF REF TO zcl_excel_style,
        lt_style_issues TYPE TABLE OF string,
        lo_iterator TYPE REF TO zcl_excel_collection_iterator.

  " Collect all used styles
  lo_iterator = lo_excel->get_styles_iterator( ).
  WHILE lo_iterator->has_next( ) = abap_true.
    DATA(lo_style) = CAST zcl_excel_style( lo_iterator->get_next( ) ).
    APPEND lo_style TO lt_used_styles.
  ENDWHILE.

  " Check for style consistency issues
  LOOP AT lt_used_styles INTO lo_style.
    " Check font consistency
    IF lo_style->font->name <> 'Arial' AND lo_style->font->name <> 'Calibri'.
      APPEND |Non-standard font used: { lo_style->font->name }| TO lt_style_issues.
    ENDIF.

    " Check color consistency
    IF lo_style->font->color->get_rgb( ) NOT IN lr_approved_colors.
      APPEND |Non-approved color used: { lo_style->font->color->get_rgb( ) }| TO lt_style_issues.
    ENDIF.

    " Check size consistency
    IF lo_style->font->size < 8 OR lo_style->font->size > 18.
      APPEND |Font size out of range: { lo_style->font->size }| TO lt_style_issues.
    ENDIF.
  ENDLOOP.

  " Report issues
  IF lt_style_issues IS NOT INITIAL.
    LOOP AT lt_style_issues INTO DATA(lv_issue).
      MESSAGE lv_issue TYPE 'W'.
    ENDLOOP.
  ENDIF.
ENDMETHOD.
```

## Complete Advanced Styling Example

### Enterprise Dashboard with Advanced Styles

```abap
" Complete example: Create enterprise dashboard with advanced styling
METHOD create_advanced_dashboard.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_style_factory TYPE REF TO zcl_excel_style_factory.

  " Initialize workbook with custom theme
  CREATE OBJECT lo_excel.
  create_custom_theme( lo_excel ).
  
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Executive Dashboard' ).

  " Create style factory for consistent styling
  CREATE OBJECT lo_style_factory
    EXPORTING
      io_excel = lo_excel.

  " Apply advanced header styling
  DATA(lo_title_style) = lo_style_factory->create_corporate_header( ).
  lo_worksheet->set_cell(
    ip_column = 'A'
    ip_row = 1
    ip_value = 'Q4 2023 Performance Dashboard'
    ip_style = lo_title_style
  ).
  lo_worksheet->set_merge( ip_range = 'A1:H1' ).

  " Create KPI section with conditional formatting
  create_kpi_section( lo_worksheet ).
  
  " Add data table with alternating row styles
  create_styled_data_table( lo_worksheet ).
  
  " Apply advanced conditional formatting
  create_advanced_conditional_formatting( lo_worksheet ).

  " Generate final file
  DATA: lo_writer TYPE REF TO zcl_excel_writer_2007.
  CREATE OBJECT lo_writer.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
ENDMETHOD.
```

## Best Practices for Advanced Styling

### Style Management Guidelines

1. **Consistency**: Use style factories and templates for consistent appearance
2. **Performance**: Cache frequently used styles to avoid duplication
3. **Maintainability**: Create hierarchical style systems for easy updates
4. **Accessibility**: Ensure sufficient contrast and readable fonts

### Advanced Techniques

1. **Theme Integration**: Leverage Excel themes for professional appearance
2. **Conditional Logic**: Use data-driven styling for dynamic presentations
3. **Responsive Design**: Adapt styles based on content characteristics
4. **Quality Control**: Implement validation for style consistency

## Next Steps

After mastering advanced custom styles:

- **[Templates](/advanced/templates)** - Create reusable styled templates
- **[Automation](/advanced/automation)** - Automate style application processes
- **[Performance Tuning](/advanced/performance-tuning)** - Optimize style-heavy workbooks
- **[Integration Patterns](/advanced/integration)** - Integrate with corporate design systems

## Common Advanced Styling Patterns

### Quick Reference for Advanced Operations

```abap
" Theme-based styling
lo_style->font->color->set_theme( zcl_excel_style_color=>c_theme_accent1 ).

" Conditional style creation
lo_style = get_or_create_performance_style( 
  iv_color = lv_dynamic_color
  iv_rating = ls_data-rating
).

" Bulk style application
lo_worksheet->set_styles_bulk( lt_style_ranges ).
