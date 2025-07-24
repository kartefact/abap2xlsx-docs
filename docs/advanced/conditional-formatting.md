# Conditional Formatting

Advanced guide to implementing conditional formatting rules in Excel worksheets with abap2xlsx.

## Understanding Conditional Formatting

Conditional formatting allows you to automatically apply formatting to cells based on their values or formulas. This creates dynamic visual representations that highlight important data patterns and trends.

## Basic Conditional Formatting

### Cell Value Rules

```abap
" Create basic conditional formatting based on cell values
METHOD create_basic_conditional_formatting.
  DATA: lo_cond_format TYPE REF TO zcl_excel_style_cond.

  " Highlight cells greater than 100
  lo_cond_format = lo_worksheet->add_new_style_cond( ).
  lo_cond_format->set_range( 'B2:B100' ).
  lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_cond_format->set_operator( zcl_excel_style_cond=>c_operator_greaterthan ).
  lo_cond_format->set_formula( '100' ).
  lo_cond_format->set_color( zcl_excel_style_color=>c_green ).
  lo_cond_format->set_font_bold( abap_true ).

  " Highlight negative values in red
  DATA(lo_negative_format) = lo_worksheet->add_new_style_cond( ).
  lo_negative_format->set_range( 'C2:C100' ).
  lo_negative_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_negative_format->set_operator( zcl_excel_style_cond=>c_operator_lessthan ).
  lo_negative_format->set_formula( '0' ).
  lo_negative_format->set_fill_color( 'FFCCCC' ).
  lo_negative_format->set_font_color( 'CC0000' ).
ENDMETHOD.
```

### Text-Based Rules

```abap
" Apply formatting based on text content
METHOD create_text_based_formatting.
  DATA: lo_text_format TYPE REF TO zcl_excel_style_cond.

  " Highlight cells containing specific text
  lo_text_format = lo_worksheet->add_new_style_cond( ).
  lo_text_format->set_range( 'A2:A100' ).
  lo_text_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_text_format->set_operator( zcl_excel_style_cond=>c_operator_containstext ).
  lo_text_format->set_formula( 'URGENT' ).
  lo_text_format->set_fill_color( 'FF0000' ).
  lo_text_format->set_font_color( 'FFFFFF' ).

  " Format cells beginning with specific text
  DATA(lo_begins_format) = lo_worksheet->add_new_style_cond( ).
  lo_begins_format->set_range( 'D2:D100' ).
  lo_begins_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_begins_format->set_operator( zcl_excel_style_cond=>c_operator_beginswith ).
  lo_begins_format->set_formula( 'COMPLETE' ).
  lo_begins_format->set_fill_color( '00FF00' ).
ENDMETHOD.
```

## Advanced Conditional Formatting

### Formula-Based Rules

```abap
" Create complex conditional formatting using formulas
METHOD create_formula_based_formatting.
  DATA: lo_formula_format TYPE REF TO zcl_excel_style_cond.

  " Highlight entire rows based on status column
  lo_formula_format = lo_worksheet->add_new_style_cond( ).
  lo_formula_format->set_range( 'A2:F100' ).
  lo_formula_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_formula_format->set_formula( '$F2="OVERDUE"' ).
  lo_formula_format->set_fill_color( 'FFE6E6' ).

  " Highlight alternating rows
  DATA(lo_alternate_format) = lo_worksheet->add_new_style_cond( ).
  lo_alternate_format->set_range( 'A2:F100' ).
  lo_alternate_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_alternate_format->set_formula( 'MOD(ROW(),2)=0' ).
  lo_alternate_format->set_fill_color( 'F0F0F0' ).

  " Highlight duplicates
  DATA(lo_duplicate_format) = lo_worksheet->add_new_style_cond( ).
  lo_duplicate_format->set_range( 'B2:B100' ).
  lo_duplicate_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_duplicate_format->set_formula( 'COUNTIF($B$2:$B$100,B2)>1' ).
  lo_duplicate_format->set_fill_color( 'FFFF99' ).
ENDMETHOD.
```

### Data Bars and Icon Sets

```abap
" Create visual indicators with data bars and icons
METHOD create_visual_indicators.
  DATA: lo_databar_format TYPE REF TO zcl_excel_style_cond,
        lo_iconset_format TYPE REF TO zcl_excel_style_cond.

  " Add data bars for numerical values
  lo_databar_format = lo_worksheet->add_new_style_cond( ).
  lo_databar_format->set_range( 'C2:C100' ).
  lo_databar_format->set_rule_type( zcl_excel_style_cond=>c_rule_databar ).
  lo_databar_format->set_color( zcl_excel_style_color=>c_blue ).
  lo_databar_format->set_show_value( abap_true ).

  " Add icon sets for performance indicators
  lo_iconset_format = lo_worksheet->add_new_style_cond( ).
  lo_iconset_format->set_range( 'D2:D100' ).
  lo_iconset_format->set_rule_type( zcl_excel_style_cond=>c_rule_iconset ).
  lo_iconset_format->set_iconset_type( zcl_excel_style_cond=>c_iconset_3trafficlights ).
  
  " Configure thresholds for icon sets
  lo_iconset_format->set_threshold( 
    iv_position = 1
    iv_type = 'percent'
    iv_value = '33'
  ).
  lo_iconset_format->set_threshold(
    iv_position = 2
    iv_type = 'percent'
    iv_value = '67'
  ).
ENDMETHOD.
```

### Color Scales

```abap
" Apply color scales for heat map visualization
METHOD create_color_scales.
  DATA: lo_colorscale_format TYPE REF TO zcl_excel_style_cond.

  " Two-color scale (red to green)
  lo_colorscale_format = lo_worksheet->add_new_style_cond( ).
  lo_colorscale_format->set_range( 'E2:E100' ).
  lo_colorscale_format->set_rule_type( zcl_excel_style_cond=>c_rule_colorscale ).
  lo_colorscale_format->set_colorscale_type( '2' ).
  lo_colorscale_format->set_min_color( 'FF0000' ).  " Red for minimum
  lo_colorscale_format->set_max_color( '00FF00' ).  " Green for maximum

  " Three-color scale (red-yellow-green)
  DATA(lo_three_color_format) = lo_worksheet->add_new_style_cond( ).
  lo_three_color_format->set_range( 'F2:F100' ).
  lo_three_color_format->set_rule_type( zcl_excel_style_cond=>c_rule_colorscale ).
  lo_three_color_format->set_colorscale_type( '3' ).
  lo_three_color_format->set_min_color( 'FF0000' ).    " Red
  lo_three_color_format->set_mid_color( 'FFFF00' ).    " Yellow
  lo_three_color_format->set_max_color( '00FF00' ).    " Green
ENDMETHOD.
```

## Top/Bottom Rules

### Ranking-Based Formatting

```abap
" Highlight top and bottom performers
METHOD create_ranking_formatting.
  DATA: lo_top_format TYPE REF TO zcl_excel_style_cond,
        lo_bottom_format TYPE REF TO zcl_excel_style_cond.

  " Highlight top 10 values
  lo_top_format = lo_worksheet->add_new_style_cond( ).
  lo_top_format->set_range( 'G2:G100' ).
  lo_top_format->set_rule_type( zcl_excel_style_cond=>c_rule_top10 ).
  lo_top_format->set_rank( 10 ).
  lo_top_format->set_bottom( abap_false ).
  lo_top_format->set_fill_color( '00FF00' ).

  " Highlight bottom 5 values
  lo_bottom_format = lo_worksheet->add_new_style_cond( ).
  lo_bottom_format->set_range( 'G2:G100' ).
  lo_bottom_format->set_rule_type( zcl_excel_style_cond=>c_rule_top10 ).
  lo_bottom_format->set_rank( 5 ).
  lo_bottom_format->set_bottom( abap_true ).
  lo_bottom_format->set_fill_color( 'FF0000' ).

  " Highlight top 20 percent
  DATA(lo_percent_format) = lo_worksheet->add_new_style_cond( ).
  lo_percent_format->set_range( 'H2:H100' ).
  lo_percent_format->set_rule_type( zcl_excel_style_cond=>c_rule_top10 ).
  lo_percent_format->set_rank( 20 ).
  lo_percent_format->set_percent( abap_true ).
  lo_percent_format->set_fill_color( 'CCFFCC' ).
ENDMETHOD.
```

## Date-Based Conditional Formatting

### Time-Sensitive Rules

```abap
" Apply formatting based on dates
METHOD create_date_based_formatting.
  DATA: lo_date_format TYPE REF TO zcl_excel_style_cond.

  " Highlight overdue dates
  lo_date_format = lo_worksheet->add_new_style_cond( ).
  lo_date_format->set_range( 'I2:I100' ).
  lo_date_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_date_format->set_operator( zcl_excel_style_cond=>c_operator_lessthan ).
  lo_date_format->set_formula( 'TODAY()' ).
  lo_date_format->set_fill_color( 'FF9999' ).

  " Highlight dates in next 7 days
  DATA(lo_upcoming_format) = lo_worksheet->add_new_style_cond( ).
  lo_upcoming_format->set_range( 'I2:I100' ).
  lo_upcoming_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_upcoming_format->set_formula( 'AND(I2>=TODAY(),I2<=TODAY()+7)' ).
  lo_upcoming_format->set_fill_color( 'FFFFCC' ).

  " Highlight this month's dates
  DATA(lo_month_format) = lo_worksheet->add_new_style_cond( ).
  lo_month_format->set_range( 'I2:I100' ).
  lo_month_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
  lo_month_format->set_formula( 'MONTH(I2)=MONTH(TODAY())' ).
  lo_month_format->set_fill_color( 'E6F3FF' ).
ENDMETHOD.
```

## Managing Multiple Rules

### Rule Priority and Interaction

```abap
" Manage multiple conditional formatting rules
METHOD manage_multiple_rules.
  DATA: lo_rule1 TYPE REF TO zcl_excel_style_cond,
        lo_rule2 TYPE REF TO zcl_excel_style_cond,
        lo_rule3 TYPE REF TO zcl_excel_style_cond.

  " Rule 1: High priority - Critical values
  lo_rule1 = lo_worksheet->add_new_style_cond( ).
  lo_rule1->set_range( 'J2:J100' ).
  lo_rule1->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_rule1->set_operator( zcl_excel_style_cond=>c_operator_greaterthan ).
  lo_rule1->set_formula( '1000' ).
  lo_rule1->set_fill_color( 'FF0000' ).
  lo_rule1->set_priority( 1 ).
  lo_rule1->set_stop_if_true( abap_true ).

  " Rule 2: Medium priority - Warning values
  lo_rule2 = lo_worksheet->add_new_style_cond( ).
  lo_rule2->set_range( 'J2:J100' ).
  lo_rule2->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_rule2->set_operator( zcl_excel_style_cond=>c_operator_greaterthan ).
  lo_rule2->set_formula( '500' ).
  lo_rule2->set_fill_color( 'FFFF00' ).
  lo_rule2->set_priority( 2 ).

  " Rule 3: Low priority - Normal values
  lo_rule3 = lo_worksheet->add_new_style_cond( ).
  lo_rule3->set_range( 'J2:J100' ).
  lo_rule3->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
  lo_rule3->set_operator( zcl_excel_style_cond=>c_operator_greaterthan ).
  lo_rule3->set_formula( '0' ).
  lo_rule3->set_fill_color( '00FF00' ).
  lo_rule3->set_priority( 3 ).
ENDMETHOD.
```

I'll continue from where the conditional formatting guide left off. Here's the completion of that file:

## `docs/advanced/conditional-formatting.md` (continued)

```markdown
## Complete Conditional Formatting Example

### Dashboard with Multiple Formatting Rules

```abap
" Complete example: Sales dashboard with conditional formatting
METHOD create_formatted_dashboard.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Dashboard' ).

  " Add sample data
  populate_dashboard_data( lo_worksheet ).

  " Apply multiple conditional formatting rules
  create_performance_indicators( lo_worksheet ).
  create_trend_analysis( lo_worksheet ).
  create_alert_system( lo_worksheet ).

  " Generate final file
  DATA: lo_writer TYPE REF TO zcl_excel_writer_2007.
  CREATE OBJECT lo_writer.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
ENDMETHOD.

METHOD create_performance_indicators.
  " Sales performance with traffic light system
  DATA(lo_performance) = lo_worksheet->add_new_style_cond( ).
  lo_performance->set_range( 'D2:D50' ).
  lo_performance->set_rule_type( zcl_excel_style_cond=>c_rule_iconset ).
  lo_performance->set_iconset_type( zcl_excel_style_cond=>c_iconset_3trafficlights ).
ENDMETHOD.
```

## Integration with Data Tables

The conditional formatting system integrates seamlessly with table binding operations <cite>src/zcl_excel_worksheet.clas.abap:1256-1264</cite>:

```abap
" Apply conditional formatting during table binding
METHOD bind_table_with_formatting.
  " Bind table data first
  lo_worksheet->bind_table(
    ip_table = lt_sales_data
    it_field_catalog = lt_field_catalog
  ).

  " Apply conditional formatting to specific columns
  LOOP AT lt_field_catalog INTO DATA(ls_field_catalog) WHERE style_cond IS NOT INITIAL.
    DATA(lo_style_cond) = lo_worksheet->get_style_cond( ls_field_catalog-style_cond ).
    lo_style_cond->set_range( 
      ip_start_column = ls_field_catalog-column_name
      ip_start_row = 2
      ip_stop_column = ls_field_catalog-column_name
      ip_stop_row = 100
    ).
  ENDLOOP.
ENDMETHOD.
```

## Writer Integration

The conditional formatting rules are processed by the Excel writer classes during file generation <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:874-1279</cite>. The writer handles different rule types including:

- **Cell value rules** (`c_rule_cellis`) <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:1141-1163</cite>
- **Data bars** (`c_rule_databar`) <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:949-987</cite>
- **Color scales** (`c_rule_colorscale`) <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:989-1046</cite>
- **Icon sets** (`c_rule_iconset`) <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:1047-1139</cite>

## Best Practices

### Performance Guidelines

1. **Limit Rules**: Avoid excessive conditional formatting rules on large ranges
2. **Optimize Formulas**: Use efficient Excel formulas in expression-based rules
3. **Range Management**: Apply rules to specific ranges rather than entire columns

### Design Guidelines

1. **Color Consistency**: Use consistent color schemes across your workbook
2. **Visual Hierarchy**: Prioritize rules to avoid conflicting formats
3. **Accessibility**: Ensure sufficient contrast for readability

## Next Steps

After mastering conditional formatting:

- **[Pivot Tables](/advanced/pivot-tables)** - Create dynamic data analysis
- **[Data Validation](/advanced/data-validation)** - Add input validation rules
- **[Password Protection](/advanced/password-protection)** - Secure your workbooks

## Common Conditional Formatting Patterns

### Quick Reference

```abap
" Basic cell value rule
lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
lo_cond_format->set_operator( zcl_excel_style_cond=>c_operator_greaterthan ).

" Formula-based rule
lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_expression ).
lo_cond_format->set_formula( 'MOD(ROW(),2)=0' ).

" Data bars
lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_databar ).
```

This guide covers the comprehensive conditional formatting capabilities of abap2xlsx, enabling you to create dynamic, visually appealing Excel reports that automatically highlight important data patterns.
