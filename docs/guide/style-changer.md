# Style Changer

`zcl_excel_style_changer` provides **range-based style operations** — applying or modifying
styles across multiple cells, rows, columns, or rectangular regions in a single call. It is
the right tool when you need to apply uniform formatting to a variable-size range determined
at runtime, without iterating cell by cell.

## Getting the Style Changer

The style changer is accessed as an attribute of the worksheet:

```abap
DATA(lo_sc) = lo_worksheet->get_style_changer( ).
```

Alternatively, create an instance and pass the worksheet:

```abap
DATA(lo_sc) = NEW zcl_excel_style_changer( io_worksheet = lo_worksheet ).
```

## Applying a Style to a Range

```abap
" Create the style to apply
DATA(lo_style) = lo_excel->add_new_style( ).
lo_style->font->bold  = abap_true.
lo_style->font->color->set_rgb( '1F4E79' ).  " Dark blue
lo_style->fill->fgcolor->set_rgb( 'DEEAF1' ).  " Light blue background
lo_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.

" Apply to a rectangular range
lo_sc->set_style_to_area(
  is_area = VALUE zexcel_s_cell_style_area(
    row_from    = 1
    row_to      = 1
    column_from = 1
    column_to   = 5
  )
  io_style = lo_style
).
```

## Setting Individual Style Properties on a Range

For targeted changes (e.g., bold all cells in a column without changing other properties),
use the individual setters. Each method reads existing styles and applies only the delta,
preserving unrelated properties:

```abap
" Make column A bold
lo_sc->set_bold(
  is_area = VALUE zexcel_s_cell_style_area(
    row_from    = 1
    row_to      = 9999  " large number = 'to end of data'
    column_from = 1
    column_to   = 1
  )
  iv_bold = abap_true
).

" Set background colour on a range
lo_sc->set_fill_color(
  is_area    = VALUE zexcel_s_cell_style_area(
    row_from    = 2  row_to      = 50
    column_from = 1  column_to   = 6
  )
  iv_fgcolor = 'FFF2CC'  " Light yellow
).

" Set font colour
lo_sc->set_font_color(
  is_area    = VALUE zexcel_s_cell_style_area(
    row_from = 2 row_to = 50 column_from = 1 column_to = 1
  )
  iv_color = 'C00000'  " Dark red
).

" Set number format on the Amount column
lo_sc->set_number_format(
  is_area = VALUE zexcel_s_cell_style_area(
    row_from = 2 row_to = 200 column_from = 5 column_to = 5
  )
  iv_format_string = '#,##0.00'
).
```

## Available Style Changer Methods

| Method | Description |
|---|---|
| `set_style_to_area` | Apply a complete `zcl_excel_style` object to a range |
| `set_bold` | Set or clear bold on a range |
| `set_italic` | Set or clear italic on a range |
| `set_underline` | Set underline style on a range |
| `set_font_color` | Set font colour (6-digit hex) on a range |
| `set_fill_color` | Set solid background fill colour on a range |
| `set_number_format` | Set a number format string on a range |
| `set_alignment` | Set horizontal/vertical alignment on a range |
| `set_borders` | Apply border configuration to a range |
| `set_wrap_text` | Enable/disable text wrap on a range |
| `change_style` | Callback-based style mutation — pass a closure or method reference |

## Conditional Formatting with `zcl_excel_style_cond`

For **rule-based** highlighting (highlight cells > threshold, colour scales, data bars),
use `zcl_excel_style_cond` directly on the worksheet. The style changer is for
programmatic range operations; conditional formatting is a separate mechanism:

```abap
DATA: lo_cond_style TYPE REF TO zcl_excel_style_cond.

lo_cond_style = lo_worksheet->add_new_style_cond( ).
lo_cond_style->set_range( 'C2:C100' ).
lo_cond_style->set_operator(
  zcl_excel_style_cond=>c_operator_greaterthan
).
lo_cond_style->formula1 = '10000'.

" Style to apply when the condition is true
DATA(lo_hl_style) = lo_excel->add_new_style( ).
lo_hl_style->fill->fgcolor->set_rgb( 'C6EFCE' ).  " Green
lo_hl_style->fill->filltype = zcl_excel_style_fill=>c_fill_solid.
lo_cond_style->set_style( lo_hl_style ).
```

See **[Formatting](/guide/formatting)** for the full conditional formatting reference.

## Performance Note

The style changer iterates the existing cell table internally. For very large worksheets
(100 000+ cells), applying a style to a huge range can be slow. In such cases:

- Apply the style **before** writing data by setting a column default style on
  `zcl_excel_column->set_default_style`.
- Use `zcl_excel_writer_huge_file` if you only need basic formatting.

## Next Steps

- **[Formatting](/guide/formatting)** — individual cell styles, number formats, conditional formatting
- **[Worksheets](/guide/worksheets)** — column/row default styles
- **[Data Conversion](/guide/data-conversion)** — `bind_table` with `fieldnames` for header styling
