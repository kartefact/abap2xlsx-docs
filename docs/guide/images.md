# Images and Drawings

Comprehensive guide to adding images, drawings, and other visual elements to Excel worksheets with abap2xlsx.

## Drawing Architecture

Images and drawings in abap2xlsx are managed through the drawing system, where visual elements are treated as drawing objects that can be positioned and sized within worksheets.

```abap
" Basic image insertion
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_drawing TYPE REF TO zcl_excel_drawing,
      lv_image_data TYPE xstring.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).

" Create drawing object
lo_drawing = lo_excel->add_new_drawing( ).
lo_drawing->set_type( zcl_excel_drawing=>type_image ).

" Set image data and properties
lo_drawing->set_media( 
  ip_media = lv_image_data 
  ip_media_type = 'image/png' 
).

" Position the image
lo_drawing->set_position(
  ip_from_row = 2
  ip_from_col = 'B'
  ip_to_row = 10
  ip_to_col = 'F'
).

" Add to worksheet
lo_worksheet->add_drawing( lo_drawing ).
```

## Image Sources

### Loading from XSTRING

```abap
" Load image from binary data
METHOD load_image_from_xstring.
  DATA: lv_image_xstring TYPE xstring,
        lo_drawing TYPE REF TO zcl_excel_drawing.

  " Your method to get image data
  lv_image_xstring = get_image_binary_data( ).

  " Create drawing with image data
  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_type( zcl_excel_drawing=>type_image ).
  lo_drawing->set_media(
    ip_media = lv_image_xstring
    ip_media_type = 'image/jpeg'
  ).

  RETURN lo_drawing.
ENDMETHOD.
```

### Loading from MIME Repository

```abap
" Load image from MIME repository
METHOD load_image_from_mime.
  DATA: lo_drawing TYPE REF TO zcl_excel_drawing.

  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_type( zcl_excel_drawing=>type_image ).
  
  " Load from MIME repository
  lo_drawing->set_media_mime(
    ip_mime_name = 'ZCOMPANY_LOGO'
    ip_media_type = 'image/png'
  ).

  RETURN lo_drawing.
ENDMETHOD.
```

### Loading from WWW Repository (SMW0)

```abap
" Load image from WWW repository
METHOD load_image_from_www.
  DATA: lo_drawing TYPE REF TO zcl_excel_drawing.

  lo_drawing = lo_excel->add_new_drawing( ).
  lo_drawing->set_type( zcl_excel_drawing=>type_image ).
  
  " Load from WWW repository
  lo_drawing->set_media_www(
    ip_key = 'ZLOGO_PNG'
    ip_media_type = 'image/png'
  ).

  RETURN lo_drawing.
ENDMETHOD.
```

## Image Positioning and Sizing

### Absolute Positioning

```abap
" Position image using cell coordinates
lo_drawing->set_position(
  ip_from_row = 1      " Start row
  ip_from_col = 'A'    " Start column
  ip_to_row = 8        " End row
  ip_to_col = 'D'      " End column
).

" Alternative: Position with offsets
lo_drawing->set_position(
  ip_from_row = 1
  ip_from_col = 'A'
  ip_from_row_offset = 5    " Pixels from top of cell
  ip_from_col_offset = 10   " Pixels from left of cell
  ip_to_row = 8
  ip_to_col = 'D'
  ip_to_row_offset = -5     " Pixels from bottom of cell
  ip_to_col_offset = -10    " Pixels from right of cell
).
```

### Pixel-Based Positioning

```abap
" Position using pixel coordinates
METHOD position_image_pixels.
  DATA: lv_x_pixels TYPE i VALUE 100,
        lv_y_pixels TYPE i VALUE 50,
        lv_width_pixels TYPE i VALUE 200,
        lv_height_pixels TYPE i VALUE 150.

  " Convert pixels to EMU (English Metric Units)
  DATA: lv_x_emu TYPE i,
        lv_y_emu TYPE i,
        lv_width_emu TYPE i,
        lv_height_emu TYPE i.

  lv_x_emu = lo_drawing->pixel2emu( lv_x_pixels ).
  lv_y_emu = lo_drawing->pixel2emu( lv_y_pixels ).
  lv_width_emu = lo_drawing->pixel2emu( lv_width_pixels ).
  lv_height_emu = lo_drawing->pixel2emu( lv_height_pixels ).

  " Set position in EMU
  lo_drawing->set_position_emu(
    ip_x = lv_x_emu
    ip_y = lv_y_emu
    ip_width = lv_width_emu
    ip_height = lv_height_emu
  ).
ENDMETHOD.
```

## Image Types and Formats

### Supported Image Formats

```abap
" Different image formats
CONSTANTS: c_image_png TYPE string VALUE 'image/png',
           c_image_jpeg TYPE string VALUE 'image/jpeg',
           c_image_gif TYPE string VALUE 'image/gif',
           c_image_bmp TYPE string VALUE 'image/bmp',
           c_image_tiff TYPE string VALUE 'image/tiff'.

" Set appropriate media type
CASE lv_file_extension.
  WHEN 'PNG' OR 'png'.
    lv_media_type = c_image_png.
  WHEN 'JPG' OR 'JPEG' OR 'jpg' OR 'jpeg'.
    lv_media_type = c_image_jpeg.
  WHEN 'GIF' OR 'gif'.
    lv_media_type = c_image_gif.
  WHEN 'BMP' OR 'bmp'.
    lv_media_type = c_image_bmp.
  WHEN OTHERS.
    lv_media_type = c_image_png.  " Default
ENDCASE.

lo_drawing->set_media(
  ip_media = lv_image_data
  ip_media_type = lv_media_type
).
```

## Advanced Image Features

### Image Scaling and Aspect Ratio

```abap
" Maintain aspect ratio while scaling
METHOD scale_image_proportionally.
  DATA: lv_original_width TYPE i VALUE 400,
        lv_original_height TYPE i VALUE 300,
        lv_target_width TYPE i VALUE 200,
        lv_target_height TYPE i,
        lv_scale_factor TYPE f.

  " Calculate scale factor
  lv_scale_factor = lv_target_width / lv_original_width.
  lv_target_height = lv_original_height * lv_scale_factor.

  " Position with calculated dimensions
  lo_drawing->set_position(
    ip_from_row = 2
    ip_from_col = 'B'
    ip_to_row = 2 + ( lv_target_height / 20 )  " Approximate row height
    ip_to_col = zcl_excel_common=>convert_column2alpha( 
      zcl_excel_common=>convert_column2int( 'B' ) + ( lv_target_width / 64 )  " Approximate column width
    )
  ).
ENDMETHOD.
```

### Image Transparency and Effects

```abap
" Configure image display properties
METHOD configure_image_effects.
  " Set image transparency (if supported by format)
  lo_drawing->set_transparency( 50 ).  " 50% transparent

  " Configure image rotation
  lo_drawing->set_rotation( 15 ).  " 15 degrees

  " Set image border
  lo_drawing->set_border(
    ip_style = 'solid'
    ip_width = 2
    ip_color = '000000'
  ).
ENDMETHOD.
```

## Working with Multiple Images

### Image Gallery Creation

```abap
" Create a gallery of images
METHOD create_image_gallery.
  DATA: lt_images TYPE TABLE OF string,
        lv_row TYPE i VALUE 2,
        lv_col TYPE string VALUE 'B',
        lv_images_per_row TYPE i VALUE 3,
        lv_current_image TYPE i VALUE 0.

  " Fill image list
  APPEND 'IMAGE1_PNG' TO lt_images.
  APPEND 'IMAGE2_PNG' TO lt_images.
  APPEND 'IMAGE3_PNG' TO lt_images.
  APPEND 'IMAGE4_PNG' TO lt_images.

  LOOP AT lt_images INTO DATA(lv_image_key).
    ADD 1 TO lv_current_image.

    " Create drawing for each image
    DATA(lo_drawing) = lo_excel->add_new_drawing( ).
    lo_drawing->set_type( zcl_excel_drawing=>type_image ).
    lo_drawing->set_media_www(
      ip_key = lv_image_key
      ip_media_type = 'image/png'
    ).

    " Position in grid layout
    lo_drawing->set_position(
      ip_from_row = lv_row
      ip_from_col = lv_col
      ip_to_row = lv_row + 6
      ip_to_col = zcl_excel_common=>convert_column2alpha( 
        zcl_excel_common=>convert_column2int( lv_col ) + 2 
      )
    ).

    lo_worksheet->add_drawing( lo_drawing ).

    " Move to next position
    IF lv_current_image MOD lv_images_per_row = 0.
      " New row
      ADD 8 TO lv_row.
      lv_col = 'B'.
    ELSE.
      " Next column
      lv_col = zcl_excel_common=>convert_column2alpha( 
        zcl_excel_common=>convert_column2int( lv_col ) + 4 
      ).
    ENDIF.
  ENDLOOP.
ENDMETHOD.
```

## Headers and Footers with Images

### Adding Images to Headers/Footers

```abap
" Add company logo to header
METHOD add_header_logo.
  DATA: lo_header_footer TYPE REF TO zcl_excel_header_footer,
        lo_header_image TYPE REF TO zcl_excel_drawing.

  lo_header_footer = lo_worksheet->get_header_footer( ).

  " Create image for header
  lo_header_image = lo_excel->add_new_drawing( ).
  lo_header_image->set_type( zcl_excel_drawing=>type_image ).
  lo_header_image->set_media_www(
    ip_key = 'COMPANY_LOGO'
    ip_media_type = 'image/png'
  ).

  " Set header with image
  lo_header_footer->set_odd_header_image(
    ip_position = 'L'  " Left position
    io_drawing = lo_header_image
  ).

  " Add text alongside image
  lo_header_footer->set_odd_header(
    '&L&G&C&"Arial,Bold"Monthly Report&R&D'
  ).
  " &G = Image placeholder, &C = Center, &R = Right, &D = Date
ENDMETHOD.
```

## Background Images

### Worksheet Background

```abap
" Set worksheet background image
METHOD set_worksheet_background.
  DATA: lv_background_image TYPE xstring.

  " Load background image data
  lv_background_image = load_background_image_data( ).

  " Set as worksheet background
  lo_worksheet->set_background_image( lv_background_image ).
ENDMETHOD.
```

## Performance Considerations

### Image Optimization

```abap
" Optimize images for Excel files
METHOD optimize_images.
  " 1. Compress images before adding
  " Use appropriate image formats (PNG for graphics, JPEG for photos)
  
  " 2. Limit image resolution
  " Resize images to appropriate dimensions before adding
  
  " 3. Avoid too many large images
  DATA: lv_max_images TYPE i VALUE 10,
        lv_max_size_kb TYPE i VALUE 500.
  
  " 4. Use image caching for repeated images
  IF mo_cached_logo IS NOT BOUND.
    mo_cached_logo = create_logo_drawing( ).
  ENDIF.
  
  " Reuse the same drawing object
  lo_worksheet->add_drawing( mo_cached_logo ).
ENDMETHOD.
```

### Memory Management

```abap
" Proper cleanup for drawing objects
METHOD cleanup_drawings.
  " Clear drawing references
  CLEAR: lo_drawing, lo_header_image.
  
  " Clear drawing collections
  DATA: lo_drawings TYPE REF TO zcl_excel_drawings.
  lo_drawings = lo_worksheet->get_drawings( ).
  IF lo_drawings IS BOUND.
    lo_drawings->clear( ).
  ENDIF.
ENDMETHOD.
```

## Complete Image Example

### Report with Logo and Charts

```abap
" Complete example: Professional report with images
METHOD create_report_with_images.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zif_excel_writer.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Company Report' ).

  " Add company logo
  DATA(lo_logo) = lo_excel->add_new_drawing( ).
  lo_logo->set_type( zcl_excel_drawing=>type_image ).
  lo_logo->set_media_www(
    ip_key = 'COMPANY_LOGO'
    ip_media_type = 'image/png'
  ).
  lo_logo->set_position(
    ip_from_row = 1
    ip_from_col = 'A'
    ip_to_row = 4
    ip_to_col = 'C'
  ).
  lo_worksheet->add_drawing( lo_logo ).

  " Add report title
  lo_worksheet->set_cell(
    ip_column = 'D'
    ip_row = 2
    ip_value = 'Annual Financial Report 2023'
  ).

  " Add data and other content
  populate_report_data( lo_worksheet ).

  " Add footer image
  add_footer_watermark( lo_worksheet ).

  " Generate Excel file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
ENDMETHOD.
```

## Next Steps

After mastering images and drawings:

- **[Data Conversion](/guide/data-conversion)** - Converting ABAP data structures to Excel
- **[ALV Integration](/guide/alv-integration)** - Converting ALV grids to Excel format
- **[Performance Optimization](/guide/performance)** - Optimize workbooks with multiple images
- **[Advanced Features](/advanced/custom-styles)** - Create sophisticated visual layouts

## Common Image Patterns

### Quick Reference for Image Operations

```abap
" Create and add basic image
DATA(lo_drawing) = lo_excel->add_new_drawing( ).
lo_drawing->set_type( zcl_excel_drawing=>type_image ).
lo_drawing->set_media( 
  ip_media = lv_image_data 
  ip_media_type = 'image/png' 
).

" Position image
lo_drawing->set_position(
  ip_from_row = 2
  ip_from_col = 'B'
  ip_to_row = 10
  ip_to_col = 'F'
).

" Add to worksheet
lo_worksheet->add_drawing( lo_drawing ).
```

This guide covers the comprehensive image and drawing capabilities of abap2xlsx. The drawing system provides extensive support for adding visual elements that enhance your Excel reports and make them more professional and engaging.
