# Working with Worksheets

Comprehensive guide to managing multiple worksheets and worksheet-specific features in abap2xlsx.

## Understanding Worksheets

### Worksheet Basics

In Excel, a workbook contains one or more worksheets. Each worksheet is represented by the `zcl_excel_worksheet` class in abap2xlsx .

```abap
" Basic worksheet operations
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

CREATE OBJECT lo_excel.

" Get the default worksheet (automatically created)
lo_worksheet = lo_excel->get_active_worksheet( ).

" Set worksheet title
lo_worksheet->set_title( 'Sales Data' ).
```

### Worksheet Properties

```abap
" Configure worksheet properties
lo_worksheet->set_title( 'Q1 Sales Report' ).

" Set worksheet visibility
lo_worksheet->set_sheet_state( zcl_excel_worksheet=>c_sheet_state_visible ).
" Other options: c_sheet_state_hidden, c_sheet_state_very_hidden

" Set tab color
DATA(lo_tabcolor) = lo_worksheet->get_tabcolor( ).
lo_tabcolor->set_rgb( 'FF0000' ).  " Red tab
```

## Creating Multiple Worksheets

### Adding New Worksheets

```abap
" Create multiple worksheets for different data sections
DATA: lo_summary_sheet TYPE REF TO zcl_excel_worksheet,
      lo_details_sheet TYPE REF TO zcl_excel_worksheet,
      lo_charts_sheet TYPE REF TO zcl_excel_worksheet.

" Add new worksheets
lo_summary_sheet = lo_excel->add_new_worksheet( ).
lo_summary_sheet->set_title( 'Summary' ).

lo_details_sheet = lo_excel->add_new_worksheet( ).
lo_details_sheet->set_title( 'Detailed Data' ).

lo_charts_sheet = lo_excel->add_new_worksheet( ).
lo_charts_sheet->set_title( 'Charts & Analysis' ).

" Set worksheet order (optional)
lo_excel->set_active_sheet_index( 1 ).  " Make Summary active
```

### Worksheet Navigation

```abap
" Navigate between worksheets
DATA: lo_worksheets TYPE REF TO zcl_excel_worksheets,
      lv_sheet_count TYPE i.

" Get all worksheets
lo_worksheets = lo_excel->get_worksheets( ).
lv_sheet_count = lo_worksheets->size( ).

WRITE: / |Workbook contains { lv_sheet_count } worksheets|.

" Access worksheets by different methods
lo_worksheet = lo_excel->get_worksheet_by_index( 2 ).        " By index (1-based)
lo_worksheet = lo_excel->get_worksheet_by_name( 'Summary' ). " By name
lo_worksheet = lo_excel->get_active_worksheet( ).            " Active sheet

" Iterate through all worksheets
DATA: lo_iterator TYPE REF TO zcl_excel_worksheets_iterator.
lo_iterator = lo_worksheets->get_iterator( ).

WHILE lo_iterator->has_next( ) = abap_true.
  lo_worksheet = lo_iterator->get_next( ).
  WRITE: / 'Processing worksheet:', lo_worksheet->get_title( ).

  " Process each worksheet
  " Your worksheet-specific logic here
ENDWHILE.
```

## Worksheet Layout and Structure

### Page Setup and Print Settings

```abap
" Configure page setup for printing
DATA: lo_page_setup TYPE REF TO zcl_excel_sheet_setup.

lo_page_setup = lo_worksheet->get_sheet_setup( ).

" Set orientation
lo_page_setup->set_orientation( zcl_excel_sheet_setup=>c_orientation_landscape ).

" Set paper size
lo_page_setup->set_paper_size( zcl_excel_sheet_setup=>c_papersize_a4 ).

" Set margins (in inches)
lo_page_setup->set_margin_left( '0.75' ).
lo_page_setup->set_margin_right( '0.75' ).
lo_page_setup->set_margin_top( '1.0' ).
lo_page_setup->set_margin_bottom( '1.0' ).
lo_page_setup->set_margin_header( '0.5' ).
lo_page_setup->set_margin_footer( '0.5' ).

" Set print area
lo_worksheet->set_print_area( 'A1:H50' ).

" Set print titles (repeat rows/columns on each page)
lo_worksheet->set_print_title_rows( '1:2' ).    " Repeat first 2 rows
lo_worksheet->set_print_title_columns( 'A:B' ). " Repeat columns A and B
```

### Freeze Panes

```abap
" Freeze panes for better navigation
" Freeze first row and first column
lo_worksheet->freeze_panes( ip_num_rows = 1 ip_num_columns = 1 ).

" Freeze multiple rows and columns
lo_worksheet->freeze_panes( ip_num_rows = 3 ip_num_columns = 2 ).

" Split panes (alternative to freeze)
lo_worksheet->set_split_panes(
  ip_x_split = 2000  " Horizontal split position
  ip_y_split = 1000  " Vertical split position
).
```

## Column and Row Management

### Column Operations

```abap
" Set column widths
DATA: lo_column TYPE REF TO zcl_excel_column.

" Set specific column width
lo_column = lo_worksheet->get_column( 'A' ).
lo_column->set_width( 15 ).

lo_column = lo_worksheet->get_column( 'B' ).
lo_column->set_width( 25 ).

" Auto-fit column width (approximate)
lo_column->set_auto_size( abap_true ).

" Hide/show columns
lo_column->set_visible( abap_false ).  " Hide column

" Set column outline level (for grouping)
lo_column->set_outline_level( 1 ).
```

### Row Operations

```abap
" Set row heights and properties
DATA: lo_row TYPE REF TO zcl_excel_row.

" Set specific row height
lo_row = lo_worksheet->get_row( 1 ).
lo_row->set_row_height( 25 ).

" Hide/show rows
lo_row->set_visible( abap_false ).  " Hide row

" Set row outline level (for grouping)
lo_row->set_outline_level( 1 ).
```

### Row and Column Grouping

```abap
" Create collapsible groups
" Group rows 5-10
DATA: lv_row TYPE i.
DO 6 TIMES.
  lv_row = 4 + sy-index.  " Rows 5-10
  lo_row = lo_worksheet->get_row( lv_row ).
  lo_row->set_outline_level( 1 ).
ENDDO.

" Group columns C-F
DATA: lv_col_alpha TYPE string.
DATA: lv_columns TYPE TABLE OF string.
APPEND 'C' TO lv_columns.
APPEND 'D' TO lv_columns.
APPEND 'E' TO lv_columns.
APPEND 'F' TO lv_columns.

LOOP AT lv_columns INTO lv_col_alpha.
  lo_column = lo_worksheet->get_column( lv_col_alpha ).
  lo_column->set_outline_level( 1 ).
ENDLOOP.
```

## Cell Comments

### Adding Comments to Cells

```abap
" Add a comment to a cell
DATA: lo_comment TYPE REF TO zcl_excel_comment.

lo_comment = lo_worksheet->add_new_comment( ).
lo_comment->set_text( ip_value = 'This is an important note.' ).
lo_comment->set_author( 'Karthikeyan' ).
lo_comment->set_ref( ip_column = 'B' ip_row = 3 ).
```

### Comment Box Positioning — Updated API (2025-06)

> **Breaking change in Jun 2025 (PR [#1316](https://github.com/abap2xlsx/abap2xlsx/pull/1316)):** The eight individual comment box geometry attributes have been consolidated into a **single structure** `ms_box` of type `zcl_excel_comment=>ty_box`. If you were setting these attributes individually before June 2025, update your code as shown below.

**Before (pre-Jun 2025):**
```abap
" Old API — individual attributes, no longer exists
lo_comment->bottom_offset = 1.
lo_comment->bottom_row    = 7.
lo_comment->left_column   = 2.
lo_comment->left_offset   = 15.
lo_comment->right_column  = 4.
lo_comment->right_offset  = 10.
lo_comment->top_offset    = 2.
lo_comment->top_row       = 2.
```

**After (Jun 2025+):**
```abap
" New API — ms_box structure
DATA: ls_box TYPE zcl_excel_comment=>ty_box.

" Use the built-in default constant as a starting point
ls_box = zcl_excel_comment=>mc_box_default.

" Override only the fields you need
ls_box-bottom_row   = 7.
ls_box-right_column = 4.
ls_box-top_row      = 2.
ls_box-left_column  = 2.

lo_comment->ms_box = ls_box.
```

The `mc_box_default` constant provides sensible defaults for all eight fields so you only need to override the values that differ from the standard position.

### Reading Comments — Copy Semantics

> **Updated in Jun 2025 (PR [#1317](https://github.com/abap2xlsx/abap2xlsx/pull/1317)):** `get_comments()` now returns a **copy** of the internal comments collection. Modifications to the returned object do not affect the worksheet's live state.

```abap
" Read comments — returns a copy
DATA: lo_comments TYPE REF TO zcl_excel_comments,
      lo_comment  TYPE REF TO zcl_excel_comment.

lo_comments = lo_worksheet->get_comments( ).

lo_comment = lo_comments->get_comment( ip_column = 'B' ip_row = 3 ).
IF lo_comment IS BOUND.
  WRITE: / 'Comment text:', lo_comment->get_text( ).
ENDIF.
```

The copy constructor for `zcl_excel_worksheet` also now correctly carries the comments collection when a worksheet is duplicated.

## Worksheet Protection

### Protecting Worksheets

```abap
" Protect worksheet with password
lo_worksheet->set_protection(
  ip_password = 'mypassword'
  ip_sheet = abap_true
  ip_objects = abap_true
  ip_scenarios = abap_true
).

" Selective protection - allow specific operations
DATA: lo_protection TYPE REF TO zcl_excel_protection.
lo_protection = lo_worksheet->get_protection( ).

lo_protection->set_password( 'mypassword' ).
lo_protection->set_sheet( abap_true ).
lo_protection->set_format_cells( abap_false ).    " Allow formatting
lo_protection->set_format_columns( abap_false ).  " Allow column formatting
lo_protection->set_format_rows( abap_false ).     " Allow row formatting
lo_protection->set_insert_columns( abap_false ).  " Allow inserting columns
lo_protection->set_insert_rows( abap_false ).     " Allow inserting rows
```

### Cell-Level Protection

```abap
" Unlock specific cells in protected worksheet
DATA: lo_style TYPE REF TO zcl_excel_style.

" Create unlocked style
lo_style = lo_excel->add_new_style( ).
lo_style->protection->locked = abap_false.

" Apply to specific cells that should remain editable
lo_worksheet->set_cell(
  ip_column = 'B'
  ip_row = 5
  ip_value = 'Editable Cell'
  ip_style = lo_style
).
```

## Advanced Worksheet Features

### Headers and Footers

```abap
" Set header and footer
DATA: lo_header_footer TYPE REF TO zcl_excel_header_footer.

lo_header_footer = lo_worksheet->get_header_footer( ).

" Set header sections
lo_header_footer->set_odd_header(
  '&L&"Arial,Bold"Company Name&C&"Arial"Sales Report&R&D'
).

" Set footer sections
lo_header_footer->set_odd_footer(
  '&LConfidential&C&P of &N&R&T'
).

" Header/Footer codes:
" &L = Left section, &C = Center section, &R = Right section
" &D = Date, &T = Time, &P = Page number, &N = Total pages
" &"FontName,Style" = Font formatting
```

### Background Images

```abap
" Set worksheet background image
DATA: lv_image_data TYPE xstring.

" Load image data (your implementation)
" lv_image_data = load_background_image( ).

lo_worksheet->set_background_image( lv_image_data ).
```

### Worksheet Views

```abap
" Configure worksheet view settings
DATA: lo_sheet_view TYPE REF TO zcl_excel_sheet_view.

lo_sheet_view = lo_worksheet->get_sheet_view( ).

" Set zoom level
lo_sheet_view->set_zoom_scale( 125 ).  " 125% zoom

" Set view type
lo_sheet_view->set_view( zcl_excel_sheet_view=>c_view_normal ).
" Other options: c_view_page_break_preview, c_view_page_layout

" Show/hide elements
lo_sheet_view->set_show_gridlines( abap_false ).
lo_sheet_view->set_show_row_col_headers( abap_false ).
lo_sheet_view->set_show_zeros( abap_false ).
```

## Worksheet Data Organization

### Named Ranges

```abap
" Create named ranges for easier reference
DATA: lo_range TYPE REF TO zcl_excel_range.

" Define a named range
lo_range = lo_excel->add_new_range( ).
lo_range->set_name( 'SalesData' ).
lo_range->set_value( 'Summary!$A$1:$E$100' ).

" Use named range in formulas
lo_worksheet->set_cell_formula(
  ip_column = 'F'
  ip_row = 1
  ip_formula = 'SUM(SalesData)'
).
```

### Data Validation

```abap
" Add data validation to cells
DATA: lo_data_validation TYPE REF TO zcl_excel_data_validation.

lo_data_validation = lo_worksheet->add_new_data_validation( ).
lo_data_validation->set_range( 'B2:B100' ).
lo_data_validation->set_type( zcl_excel_data_validation=>c_type_list ).
lo_data_validation->set_formula1( 'North,South,East,West' ).
lo_data_validation->set_allow_blank( abap_false ).
lo_data_validation->set_show_dropdown( abap_true ).

" Set validation error message
lo_data_validation->set_error_title( 'Invalid Region' ).
lo_data_validation->set_error( 'Please select a valid region from the dropdown.' ).
```

## Performance Considerations

### Efficient Worksheet Operations

```abap
" Batch operations for better performance
METHOD populate_worksheet_efficiently.
  " 1. Minimize worksheet switches
  " Process all data for one worksheet before moving to next

  " 2. Use table binding for large datasets
  lo_worksheet->bind_table( ip_table = lt_large_data ).

  " 3. Set styles once, reuse multiple times
  DATA(lo_header_style) = lo_excel->add_new_style( ).
  " Configure style once

  " Apply to multiple cells
  LOOP AT lt_headers INTO DATA(ls_header).
    lo_worksheet->set_cell(
      ip_column = ls_header-column
      ip_row = 1
      ip_value = ls_header-text
      ip_style = lo_header_style
    ).
  ENDLOOP.

  " 4. Clear objects when done
  CLEAR: lo_worksheet, lo_header_style.
ENDMETHOD.
```

## Next Steps

After mastering worksheet management:

- **[Cell Formatting](/guide/formatting)** - Style individual cells and ranges
- **[Excel Formulas](/guide/formulas)** - Add calculations across worksheets
- **[Charts and Graphs](/guide/charts)** - Create visual representations
- **[Data Conversion](/guide/data-conversion)** - Efficiently populate worksheets with ABAP data
- **[Reading Excel](/guide/reading-excel)** - Read back and process existing workbooks
- **[Changelog](/guide/changelog)** - Full history of recent changes

## Common Worksheet Patterns

### Multi-Sheet Report Structure

```abap
" Standard pattern for multi-sheet reports
METHOD create_multi_sheet_report.
  " Sheet 1: Executive Summary
  DATA(lo_summary) = lo_excel->add_new_worksheet( ).
  lo_summary->set_title( 'Executive Summary' ).

  " Sheet 2: Detailed Data
  DATA(lo_details) = lo_excel->add_new_worksheet( ).
  lo_details->set_title( 'Detailed Data' ).

  " Sheet 3: Charts and Analysis
  DATA(lo_charts) = lo_excel->add_new_worksheet( ).
  lo_charts->set_title( 'Analysis' ).

  " Configure each sheet appropriately
  setup_summary_sheet( lo_summary ).
  setup_details_sheet( lo_details ).
  setup_charts_sheet( lo_charts ).
ENDMETHOD.
```

This guide covers the essential techniques for managing worksheets in abap2xlsx. Proper worksheet organization is key to creating professional, navigable Excel reports.
