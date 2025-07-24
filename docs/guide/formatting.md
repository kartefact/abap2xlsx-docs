# Cell Formatting

Comprehensive guide to styling and formatting Excel cells, rows, and columns with abap2xlsx.

## Style Architecture

The abap2xlsx style system is built around the `zcl_excel_style` class and its components:

```abap
" Create and configure a style
DATA: lo_style TYPE REF TO zcl_excel_style.

lo_style = lo_excel->add_new_style( ).

" Configure font properties
lo_style->font->bold = abap_true.
lo_style->font->size = 12.
lo_style->font->color->set_rgb( '0000FF' ).

" Configure fill properties
lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_style->fill->fgcolor->set_rgb( 'FFFF00' ).

" Apply to cell
lo_worksheet->set_cell( 
  ip_column = 'A' 
  ip_row = 1 
  ip_value = 'Styled Cell'
  ip_style = lo_style
).
```

## Font Formatting

### Basic Font Properties

```abap
" Font configuration options
lo_style->font->name = 'Arial'.
lo_style->font->size = 14.
lo_style->font->bold = abap_true.
lo_style->font->italic = abap_true.
lo_style->font->underline = zcl_excel_style_font=>c_underline_single.
lo_style->font->strikethrough = abap_true.
```

### Font Colors

```abap
" Set font color using RGB
lo_style->font->color->set_rgb( 'FF0000' ).  " Red

" Set font color using theme colors
lo_style->font->color->set_theme( zcl_excel_style_color=>c_theme_accent1 ).

" Set font color using indexed colors
lo_style->font->color->set_indexed( zcl_excel_style_color=>c_indexed_red ).
```

## Cell Backgrounds and Fills

### Solid Fills

```abap
" Solid background color
lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_style->fill->fgcolor->set_rgb( 'E6E6FA' ).  " Light purple
```

### Pattern Fills

```abap
" Pattern fills with two colors
lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_pattern_darkgray.
lo_style->fill->fgcolor->set_rgb( '000000' ).  " Foreground: Black
lo_style->fill->bgcolor->set_rgb( 'FFFFFF' ).  " Background: White
```

## Borders and Lines

### Individual Borders

```abap
" Configure individual borders
lo_style->borders->left->border_style = zcl_excel_style_border=>c_border_thin.
lo_style->borders->left->border_color->set_rgb( '000000' ).

lo_style->borders->right->border_style = zcl_excel_style_border=>c_border_thick.
lo_style->borders->right->border_color->set_rgb( 'FF0000' ).

lo_style->borders->top->border_style = zcl_excel_style_border=>c_border_double.
lo_style->borders->bottom->border_style = zcl_excel_style_border=>c_border_dotted.
```

### All Borders at Once

```abap
" Apply same border to all sides
CREATE OBJECT lo_style->borders->allborders.
lo_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_medium.
lo_style->borders->allborders->border_color->set_rgb( '808080' ).
```

## Text Alignment

### Horizontal Alignment

```abap
" Horizontal alignment options
lo_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_left.
lo_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
lo_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_right.
lo_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_justify.
```

### Vertical Alignment

```abap
" Vertical alignment options
lo_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_top.
lo_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
lo_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_bottom.
```

### Text Wrapping and Rotation

```abap
" Text wrapping and rotation
lo_style->alignment->wraptext = abap_true.
lo_style->alignment->textrotation = 45.  " 45 degrees
lo_style->alignment->shrinktofit = abap_true.
```

## Number Formatting

### Built-in Number Formats

```abap
" Use predefined number formats
lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.
lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_ddmmyyyy_new.
lo_style->number_format->format_code = zcl_excel_style_number_format=>c_format_percentage_00.
```

### Custom Number Formats

```abap
" Custom number format patterns
lo_style->number_format->format_code = '#,##0.00_);[Red](#,##0.00)'.  " Currency with red negatives
lo_style->number_format->format_code = '0.00%'.  " Percentage with 2 decimals
lo_style->number_format->format_code = 'dd/mm/yyyy hh:mm'.  " Date and time
```

## Style Reuse and Management

### Creating Style Templates

```abap
" Create reusable style templates
CLASS zcl_excel_style_templates DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: get_header_style
                     IMPORTING io_excel TYPE REF TO zcl_excel
                     RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style,
                   get_currency_style
                     IMPORTING io_excel TYPE REF TO zcl_excel
                     RETURNING VALUE(ro_style) TYPE REF TO zcl_excel_style.
ENDCLASS.

CLASS zcl_excel_style_templates IMPLEMENTATION.
  METHOD get_header_style.
    ro_style = io_excel->add_new_style( ).
    ro_style->font->bold = abap_true.
    ro_style->font->color->set_rgb( 'FFFFFF' ).
    ro_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
    ro_style->fill->fgcolor->set_rgb( '366092' ).
    ro_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
  ENDMETHOD.

  METHOD get_currency_style.
    ro_style = io_excel->add_new_style( ).
    ro_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.
    ro_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_right.
  ENDMETHOD.
ENDCLASS.
```

### Applying Styles to Ranges

```abap
" Apply style to entire range
lo_worksheet->set_cell_style(
  ip_range = 'A1:D10'
  ip_style = lo_style
).

" Apply different styles to different ranges
lo_worksheet->set_cell_style( ip_range = 'A1:D1' ip_style = lo_header_style ).
lo_worksheet->set_cell_style( ip_range = 'D2:D10' ip_style = lo_currency_style ).
```

## Advanced Styling Features

### Conditional Formatting

```abap
" Add conditional formatting rules
DATA: lo_cond_format TYPE REF TO zcl_excel_style_cond.

lo_cond_format = lo_worksheet->add_new_style_cond( ).
lo_cond_format->set_range( 'C2:C100' ).
lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_cellis ).
lo_cond_format->set_operator( zcl_excel_style_cond=>c_operator_greaterthan ).
lo_cond_format->set_formula( '1000' ).
lo_cond_format->set_color( zcl_excel_style_color=>c_green ).
```

### Data Bars and Color Scales

```abap
" Data bars for visual representation
lo_cond_format = lo_worksheet->add_new_style_cond( ).
lo_cond_format->set_range( 'E2:E100' ).
lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_databar ).
lo_cond_format->set_color( zcl_excel_style_color=>c_blue ).
```

## Table Formatting

### Excel Tables with Built-in Styles

```abap
" Create formatted table with built-in style
lo_worksheet->bind_table(
  ip_table = lt_data
  is_table_settings = VALUE #(
    top_left_column = 'A'
    top_left_row = 1
    table_style = zcl_excel_table=>builtinstyle_medium9
    show_row_stripes = abap_true
    show_first_column = abap_true
    show_last_column = abap_false
  )
).
```

### Custom Table Styling

```abap
" Apply custom styling to table ranges
DATA: lo_header_style TYPE REF TO zcl_excel_style,
      lo_data_style TYPE REF TO zcl_excel_style,
      lo_total_style TYPE REF TO zcl_excel_style.

" Header row style
lo_header_style = lo_excel->add_new_style( ).
lo_header_style->font->bold = abap_true.
lo_header_style->font->color->set_rgb( 'FFFFFF' ).
lo_header_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_header_style->fill->fgcolor->set_rgb( '4472C4' ).

" Data row style
lo_data_style = lo_excel->add_new_style( ).
lo_data_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.

" Total row style
lo_total_style = lo_excel->add_new_style( ).
lo_total_style->font->bold = abap_true.
lo_total_style->borders->top->border_style = zcl_excel_style_border=>c_border_double.

" Apply styles to appropriate ranges
lo_worksheet->set_cell_style( ip_range = 'A1:E1' ip_style = lo_header_style ).
lo_worksheet->set_cell_style( ip_range = 'A2:E10' ip_style = lo_data_style ).
lo_worksheet->set_cell_style( ip_range = 'A11:E11' ip_style = lo_total_style ).
```

## Performance Considerations

### Efficient Style Management

```abap
" Best practices for style performance
METHOD apply_styles_efficiently.
  " 1. Reuse styles - create once, apply multiple times
  DATA(lo_header_style) = create_header_style( lo_excel ).
  
  " 2. Apply styles to ranges rather than individual cells
  lo_worksheet->set_cell_style( 
    ip_range = 'A1:Z1' 
    ip_style = lo_header_style 
  ).
  
  " 3. Minimize style variations
  " Use a limited set of predefined styles
  
  " 4. Cache frequently used styles
  IF mo_cached_currency_style IS NOT BOUND.
    mo_cached_currency_style = create_currency_style( lo_excel ).
  ENDIF.
ENDMETHOD.
```

## Complete Formatting Example

### Professional Report Styling

```abap
" Complete example: Create professionally formatted report
METHOD create_formatted_report.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zif_excel_writer.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Report' ).

  " Create style palette
  DATA: lo_title_style TYPE REF TO zcl_excel_style,
        lo_header_style TYPE REF TO zcl_excel_style,
        lo_currency_style TYPE REF TO zcl_excel_style,
        lo_date_style TYPE REF TO zcl_excel_style,
        lo_total_style TYPE REF TO zcl_excel_style.

  " Title style (large, bold, centered)
  lo_title_style = lo_excel->add_new_style( ).
  lo_title_style->font->size = 16.
  lo_title_style->font->bold = abap_true.
  lo_title_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.

  " Header style (white text on blue background)
  lo_header_style = lo_excel->add_new_style( ).
  lo_header_style->font->bold = abap_true.
  lo_header_style->font->color->set_rgb( 'FFFFFF' ).
  lo_header_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
  lo_header_style->fill->fgcolor->set_rgb( '4472C4' ).
  lo_header_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.

  " Currency style
  lo_currency_style = lo_excel->add_new_style( ).
  lo_currency_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.

  " Date style
  lo_date_style = lo_excel->add_new_style( ).
  lo_date_style->number_format->format_code = zcl_excel_style_number_format=>c_format_date_ddmmyyyy_new.

  " Total style (bold with top border)
  lo_total_style = lo_excel->add_new_style( ).
  lo_total_style->font->bold = abap_true.
  lo_total_style->borders->top->border_style = zcl_excel_style_border=>c_border_double.
  lo_total_style->number_format->format_code = zcl_excel_style_number_format=>c_format_currency_usd_simple.

  " Apply formatting to report
  " Title
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Monthly Sales Report' ip_style = lo_title_style ).
  lo_worksheet->set_merge( ip_range = 'A1:E1' ).

  " Headers
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 3 ip_value = 'Date' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 3 ip_value = 'Product' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 3 ip_value = 'Quantity' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 3 ip_value = 'Unit Price' ip_style = lo_header_style ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 3 ip_value = 'Total' ip_style = lo_header_style ).

  " Sample data with appropriate formatting
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 4 ip_value = '20231201' ip_style = lo_date_style ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 4 ip_value = 'Laptop' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 4 ip_value = 5 ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 4 ip_value = '999.99' ip_style = lo_currency_style ).
  lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 4 ip_formula = 'C4*D4' ip_style = lo_currency_style ).

  " Total row
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 6 ip_value = 'TOTAL:' ip_style = lo_total_style ).
  lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 6 ip_formula = 'SUM(E4:E5)' ip_style = lo_total_style ).

  " Generate file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
ENDMETHOD.
```

## Best Practices for Formatting

### Style Consistency

1. **Create a Style Guide**: Define standard styles for headers, data, totals, etc.
2. **Reuse Styles**: Create styles once and apply them multiple times
3. **Limit Color Palette**: Use a consistent set of colors throughout your reports
4. **Professional Fonts**: Stick to standard fonts like Arial, Calibri, or Times New Roman

### Performance Optimization

1. **Batch Style Applications**: Apply styles to ranges rather than individual cells
2. **Minimize Style Objects**: Avoid creating unnecessary style variations
3. **Cache Frequently Used Styles**: Store commonly used styles in class attributes
4. **Use Built-in Table Styles**: Leverage Excel's built-in table styles when possible

### Accessibility Considerations

1. **High Contrast**: Ensure sufficient contrast between text and background colors
2. **Color Independence**: Don't rely solely on color to convey information
3. **Readable Fonts**: Use appropriate font sizes (minimum 10pt for body text)
4. **Clear Structure**: Use consistent formatting to establish visual hierarchy

## Next Steps

After mastering cell formatting:

- **[Excel Formulas](/guide/formulas)** - Add calculations and dynamic content
- **[Charts and Graphs](/guide/charts)** - Create visual data representations
- **[Conditional Formatting](/advanced/conditional-formatting)** - Advanced conditional styling
- **[Custom Styles](/advanced/custom-styles)** - Create sophisticated style templates

## Common Formatting Patterns

### Quick Reference for Styling Operations

```abap
" Create and apply basic styles
DATA(lo_style) = lo_excel->add_new_style( ).
lo_style->font->bold = abap_true.
lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_style->fill->fgcolor->set_rgb( 'CCCCCC' ).

" Apply to cell
lo_worksheet->set_cell( 
  ip_column = 'A' 
  ip_row = 1 
  ip_value = 'Formatted Cell'
  ip_style = lo_style 
).

" Apply to range
lo_worksheet->set_cell_style( 
  ip_range = 'A1:E1' 
  ip_style = lo_style 
).
```

This guide covers the comprehensive formatting capabilities of abap2xlsx. [1](#22-0)  The style system provides extensive control over the visual appearance of your Excel reports, enabling you to create professional, branded documents that effectively communicate your data.
