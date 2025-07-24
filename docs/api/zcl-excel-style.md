# Style and Formatting Classes

Comprehensive guide to styling and formatting Excel cells, rows, and columns.

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

## Performance Considerations

1. **Reuse Styles**: Create styles once and reuse them across multiple cells
2. **Batch Operations**: Apply styles to ranges rather than individual cells
3. **Minimize Style Objects**: Avoid creating unnecessary style variations
4. **Cache Style References**: Store frequently used styles in variables
