# Cell Comments

abap2xlsx supports Excel-style cell comments (the yellow sticky-note annotations visible when you hover over a cell). Comments are managed through `zcl_excel_comment` (one annotation) and `zcl_excel_comments` (the per-worksheet collection accessible via `zcl_excel_worksheet->comments`).

## Adding a Comment

```abap
DATA lo_comment TYPE REF TO zcl_excel_comment.

lo_comment = NEW zcl_excel_comment( ).

" Position the comment on cell C5
lo_comment->cell_column = 'C'.
lo_comment->cell_row    = 5.

" Author and text
lo_comment->author      = 'Karthik'.

" Use ms_box structure to set text (post-PR #1316)
lo_comment->ms_box-text = 'Review this value — check against PO line 3.'.

" Optional: size and position of the comment box (in EMU)
lo_comment->ms_box-margin_left   = 1200000.
lo_comment->ms_box-margin_top    = 800000.
lo_comment->ms_box-width         = 2400000.
lo_comment->ms_box-height        = 1200000.

lo_worksheet->comments->add( lo_comment ).
```

## The `ms_box` Structure

PR #1316 (June 2025) wrapped all box-geometry and text parameters into the `ms_box` component structure. The relevant fields are:

| Field | Type | Description |
|---|---|---|
| `ms_box-text` | `string` | Comment text content |
| `ms_box-margin_left` | `i` | Left edge offset from the anchor cell, in EMU |
| `ms_box-margin_top` | `i` | Top edge offset from the anchor cell, in EMU |
| `ms_box-width` | `i` | Comment box width in EMU |
| `ms_box-height` | `i` | Comment box height in EMU |

::: tip EMU conversion
1 cm = 360 000 EMU. 1 inch = 914 400 EMU. A typical comment box of ~5 cm × 2.5 cm is approximately `width = 1 800 000`, `height = 900 000`.
:::

The `author` attribute remains a direct public attribute on the object (not inside `ms_box`).

## Cell Position

`cell_column` accepts an alphabetic column identifier (`'A'`–`'XFD'`). `cell_row` accepts an integer row number starting from 1. Both are public attributes set directly on the `zcl_excel_comment` instance.

## Reading Comments Back

The reader (`zcl_excel_reader_2007`) restores comment text and position when it parses an xlsx file that contains a `comments1.xml` part. After reading, iterate the comments collection:

```abap
DATA lo_iterator TYPE REF TO zcl_excel_collection_iterator.
DATA lo_obj      TYPE REF TO object.
DATA lo_cmt      TYPE REF TO zcl_excel_comment.

lo_iterator = lo_worksheet->comments->get_iterator( ).
WHILE lo_iterator->has_next( ) = abap_true.
  lo_obj = lo_iterator->get_next( ).
  lo_cmt ?= lo_obj.
  WRITE: / lo_cmt->cell_column, lo_cmt->cell_row, lo_cmt->ms_box-text.
ENDWHILE.
```

## Worksheet Copy and Comments

When you copy a worksheet with `zcl_excel->copy_worksheet( )`, the comments collection reference is correctly passed to the new sheet (fixed in the worksheet copy constructor — part of the v6 bug-fix stream). No extra steps are needed.

## Limitations

- abap2xlsx writes **legacy VML-based** comments (`xl/drawings/vmlDrawing1.vml`), which is the format used by Excel 97–2019 `.xlsx` files. These render correctly in all desktop versions of Excel and LibreOffice Calc.
- Threaded / modern comments (the Microsoft 365 "@mention" comment style stored in `xl/threadedComments/`) are **not** supported.
- Rich-text formatting inside comment text (bold author name etc.) is not exposed — the text is written as plain text.
- Comment box fill colour and border style cannot be customised through the public API.
