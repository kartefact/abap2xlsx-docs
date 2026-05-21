# Working with Worksheets

Comprehensive guide to managing multiple worksheets and worksheet-specific features in abap2xlsx.

## Understanding Worksheets

### Worksheet Basics

In Excel, a workbook contains one or more worksheets. Each worksheet is represented by the `zcl_excel_worksheet` class in abap2xlsx.

```abap
" Basic worksheet operations
DATA: lo_excel       TYPE REF TO zcl_excel,
      lo_worksheet   TYPE REF TO zcl_excel_worksheet.

CREATE OBJECT lo_excel.

" get_active_worksheet returns the first sheet, which zcl_excel creates automatically
lo_worksheet = lo_excel->get_active_worksheet( ).

" Sheet title appears as the tab label in Excel
lo_worksheet->set_title( 'Sales Data' ).
```

### Worksheet Properties

```abap
" Configure worksheet properties
lo_worksheet->set_title( 'Q1 Sales Report' ).

" Control sheet visibility — very_hidden sheets cannot be shown via the Excel UI
lo_worksheet->set_sheet_state( zcl_excel_worksheet=>c_sheet_state_visible ).
" Other options: c_sheet_state_hidden, c_sheet_state_very_hidden

" Colour the sheet tab
DATA(lo_tabcolor) = lo_worksheet->get_tabcolor( ).
lo_tabcolor->set_rgb( 'FF0000' ).  " Red tab — use a 6-digit hex RGB string
```

## Creating Multiple Worksheets

### Adding New Worksheets

```abap
" Create multiple worksheets for different data sections
DATA: lo_summary_sheet TYPE REF TO zcl_excel_worksheet,
      lo_details_sheet TYPE REF TO zcl_excel_worksheet,
      lo_charts_sheet  TYPE REF TO zcl_excel_worksheet.

" Each add_new_worksheet call appends a tab at the right of the tab bar
lo_summary_sheet = lo_excel->add_new_worksheet( ).
lo_summary_sheet->set_title( 'Summary' ).

lo_details_sheet = lo_excel->add_new_worksheet( ).
lo_details_sheet->set_title( 'Detailed Data' ).

lo_charts_sheet = lo_excel->add_new_worksheet( ).
lo_charts_sheet->set_title( 'Charts & Analysis' ).

" Make the first tab active when the file opens (index is 1-based)
lo_excel->set_active_sheet_index( 1 ).
```

### Worksheet Navigation

```abap
" Navigate between worksheets
DATA: lo_worksheets  TYPE REF TO zcl_excel_worksheets,
      lv_sheet_count TYPE i.

" The worksheets collection mirrors the tab order
lo_worksheets  = lo_excel->get_worksheets( ).
lv_sheet_count = lo_worksheets->size( ).

WRITE: / |Workbook contains { lv_sheet_count } worksheets|.

" Three equivalent ways to retrieve a specific worksheet reference
lo_worksheet = lo_excel->get_worksheet_by_index( 2 ).        " By tab position (1-based)
lo_worksheet = lo_excel->get_worksheet_by_name( 'Summary' ). " By exact tab label
lo_worksheet = lo_excel->get_active_worksheet( ).            " Currently active tab

" Use an iterator to process every sheet in order
DATA: lo_iterator TYPE REF TO zcl_excel_worksheets_iterator.
lo_iterator = lo_worksheets->get_iterator( ).

WHILE lo_iterator->has_next( ) = abap_true.
  lo_worksheet = lo_iterator->get_next( ).
  WRITE: / 'Processing worksheet:', lo_worksheet->get_title( ).
  " Your worksheet-specific logic here
ENDWHILE.
```

## Worksheet Layout and Structure

### Page Setup and Print Settings

```abap
" Configure page setup for printing
DATA: lo_page_setup TYPE REF TO zcl_excel_sheet_setup.

lo_page_setup = lo_worksheet->get_sheet_setup( ).

" Orientation: portrait (default) or landscape
lo_page_setup->set_orientation( zcl_excel_sheet_setup=>c_orientation_landscape ).

" Paper size constant maps to the OOXML paperSize numeric attribute
lo_page_setup->set_paper_size( zcl_excel_sheet_setup=>c_papersize_a4 ).

" Margins are specified in inches (decimal string)
lo_page_setup->set_margin_left( '0.75' ).
lo_page_setup->set_margin_right( '0.75' ).
lo_page_setup->set_margin_top( '1.0' ).
lo_page_setup->set_margin_bottom( '1.0' ).
lo_page_setup->set_margin_header( '0.5' ).
lo_page_setup->set_margin_footer( '0.5' ).

" Restrict printing to a named range of cells
lo_worksheet->set_print_area( 'A1:H50' ).

" Repeat title rows and columns on every printed page
lo_worksheet->set_print_title_rows( '1:2' ).    " Repeat first 2 rows on each page
lo_worksheet->set_print_title_columns( 'A:B' ). " Repeat columns A and B on each page
```

### Page Breaks

Page breaks let you control exactly where Excel splits a worksheet for printing. The
`zcl_excel_worksheet_pagebreaks` class (accessible via `lo_worksheet->page_breaks`)
provides two methods:

#### Adding a Page Break

```abap
" Insert a horizontal page break before row 30
" (rows 1-29 print on page 1, rows 30+ continue on page 2)
lo_worksheet->page_breaks->add_pagebreak(
  ip_column = 'A'   " column position (for a row break, use any column; 'A' is conventional)
  ip_row    = 30    " break occurs BEFORE this row
).

" Insert a vertical page break before column E
lo_worksheet->page_breaks->add_pagebreak(
  ip_column = 'E'   " break occurs BEFORE this column
  ip_row    = 1     " row position (for a column break, use 1 or the header row)
).
```

> **Row vs column breaks:** Excel distinguishes row breaks (horizontal) and column breaks
> (vertical) by the combination of `ip_column` and `ip_row` values passed to the writer.
> A break at (`ip_column = 'A'`, `ip_row = N`) is serialised as a row page break.
> A break at (`ip_column = 'X'`, `ip_row = 1`) is serialised as a column page break.

#### Reading All Breaks (Diagnostic / Test Use)

```abap
" Retrieve all registered breaks as a hashed table
DATA(lt_breaks) = lo_worksheet->page_breaks->get_all_pagebreaks( ).

" Structure of each row: cell_row (zexcel_cell_row), cell_column (zexcel_cell_column)
LOOP AT lt_breaks INTO DATA(ls_break).
  WRITE: / 'Break at row:', ls_break-cell_row,
             'column:', ls_break-cell_column.
ENDLOOP.
```

#### Combining Page Breaks with Print Area

```abap
" Typical pattern: restrict print area, add section breaks
lo_page_setup->set_paper_size( zcl_excel_sheet_setup=>c_papersize_a4 ).
lo_page_setup->set_orientation( zcl_excel_sheet_setup=>c_orientation_portrait ).
lo_worksheet->set_print_area( 'A1:H100' ).

" Add breaks every 30 data rows
DATA(lv_break_row) = 31.  " Row 1 = header; rows 2-30 = first page
WHILE lv_break_row <= 100.
  lo_worksheet->page_breaks->add_pagebreak( ip_column = 'A'  ip_row = lv_break_row ).
  ADD 30 TO lv_break_row.
ENDWHILE.
```

### Freeze Panes

```abap
" Freeze panes lock rows/columns so they stay visible while scrolling

" Freeze first row (header) and first column (row labels)
lo_worksheet->freeze_panes( ip_num_rows = 1 ip_num_columns = 1 ).

" Freeze a larger area — e.g. 3 header rows and 2 label columns
lo_worksheet->freeze_panes( ip_num_rows = 3 ip_num_columns = 2 ).

" Split panes divide the view into independently-scrollable quadrants
lo_worksheet->set_split_panes(
  ip_x_split = 2000  " Horizontal split position in 1/20th of a point
  ip_y_split = 1000  " Vertical split position in 1/20th of a point
).
```

## Column and Row Management

### Column Operations

```abap
" Set column widths
DATA: lo_column TYPE REF TO zcl_excel_column.

" Width is specified in Excel's 'character units' (roughly, characters of the default font)
lo_column = lo_worksheet->get_column( 'A' ).
lo_column->set_width( 15 ).

lo_column = lo_worksheet->get_column( 'B' ).
lo_column->set_width( 25 ).

" Auto-size asks Excel to fit the column to its widest cell on next open
lo_column->set_auto_size( abap_true ).

" Hidden columns are excluded from the view but remain in the data
lo_column->set_visible( abap_false ).

" Outline level >= 1 makes the column part of a collapsible group
lo_column->set_outline_level( 1 ).
```

### Row Operations

```abap
" Set row heights and properties
DATA: lo_row TYPE REF TO zcl_excel_row.

" Row height in points (same unit as Excel's row height dialog)
lo_row = lo_worksheet->get_row( 1 ).
lo_row->set_row_height( 25 ).

" Hidden rows are excluded from the view but remain in the data
lo_row->set_visible( abap_false ).

" Outline level >= 1 makes the row part of a collapsible group
lo_row->set_outline_level( 1 ).
```

### Row and Column Grouping

```abap
" Create collapsible groups by setting outline levels

" Group rows 5-10 at level 1 (one '+' button in the margin)
DATA: lv_row TYPE i.
DO 6 TIMES.
  lv_row = 4 + sy-index.  " sy-index runs 1..6 -> rows 5..10
  lo_row = lo_worksheet->get_row( lv_row ).
  lo_row->set_outline_level( 1 ).
ENDDO.

" Group columns C-F at level 1
DATA: lv_col_alpha TYPE string.
DATA: lv_columns   TYPE TABLE OF string.
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

" add_new_comment creates and registers the comment on the worksheet
lo_comment = lo_worksheet->add_new_comment( ).
lo_comment->set_text( ip_value = 'This is an important note.' ).
lo_comment->set_author( 'Karthikeyan' ).             " Displayed as the comment author in Excel
lo_comment->set_ref( ip_column = 'B' ip_row = 3 ).  " Cell the comment is anchored to
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
" New API — all eight geometry fields are grouped in one structure
DATA: ls_box TYPE zcl_excel_comment=>ty_box.

" Start from the built-in default so you only override the fields that differ
ls_box = zcl_excel_comment=>mc_box_default.

" Override only the values you need to change from the standard position
ls_box-bottom_row   = 7.
ls_box-right_column = 4.
ls_box-top_row      = 2.
ls_box-left_column  = 2.

lo_comment->ms_box = ls_box.  " Assign the completed structure back to the comment
```

The `mc_box_default` constant provides sensible defaults for all eight fields so you only need to override the values that differ from the standard position.

### Reading Comments — Copy Semantics

> **Updated in Jun 2025 (PR [#1317](https://github.com/abap2xlsx/abap2xlsx/pull/1317)):** `get_comments()` now returns a **copy** of the internal comments collection. Modifications to the returned object do not affect the worksheet's live state.

```abap
" Read comments — get_comments() returns a snapshot copy, not a live reference
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
  ip_password  = 'mypassword'
  ip_sheet     = abap_true     " Activate sheet protection
  ip_objects   = abap_true     " Protect embedded objects
  ip_scenarios = abap_true     " Protect named scenarios
).

" Fine-grained protection — allow specific operations even on a protected sheet
DATA: lo_protection TYPE REF TO zcl_excel_protection.
lo_protection = lo_worksheet->get_protection( ).

lo_protection->set_password( 'mypassword' ).
lo_protection->set_sheet( abap_true ).
lo_protection->set_format_cells( abap_false ).    " Users may still format cells
lo_protection->set_format_columns( abap_false ).  " Users may still resize columns
lo_protection->set_format_rows( abap_false ).     " Users may still resize rows
lo_protection->set_insert_columns( abap_false ).  " Users may still insert columns
lo_protection->set_insert_rows( abap_false ).     " Users may still insert rows
```

### Cell-Level Protection

```abap
" In a protected sheet every cell is locked by default.
" Create an 'unlocked' style and apply it to the cells that must remain editable.
DATA: lo_style TYPE REF TO zcl_excel_style.

" Add a new style to the workbook's style registry
lo_style = lo_excel->add_new_style( ).
lo_style->protection->locked = abap_false.  " This is the key flag

" Apply the unlocked style to the input cells
lo_worksheet->set_cell(
  ip_column = 'B'
  ip_row    = 5
  ip_value  = 'Editable Cell'
  ip_style  = lo_style
).
```

## Advanced Worksheet Features

### Headers and Footers

```abap
" Set header and footer
DATA: lo_header_footer TYPE REF TO zcl_excel_header_footer.

lo_header_footer = lo_worksheet->get_header_footer( ).

" Header/footer code reference:
" &L = Left section    &C = Centre section    &R = Right section
" &D = Current date   &T = Current time      &P = Page number   &N = Total pages
" &"FontName,Style"   sets font name and style for the following text
lo_header_footer->set_odd_header(
  '&L&"Arial,Bold"Company Name&C&"Arial"Sales Report&R&D'
).

lo_header_footer->set_odd_footer(
  '&LConfidential&C&P of &N&R&T'
).
```

### Background Images

```abap
" Set worksheet background image (tiled fill behind the cells)
DATA: lv_image_data TYPE xstring.

" Load image data as binary XSTRING from BDS, MIME repository, or file upload
" lv_image_data = load_background_image( ).

lo_worksheet->set_background_image( lv_image_data ).
```

### Worksheet Views

```abap
" Configure worksheet view settings
DATA: lo_sheet_view TYPE REF TO zcl_excel_sheet_view.

lo_sheet_view = lo_worksheet->get_sheet_view( ).

" Zoom percentage — valid range is roughly 10-400
lo_sheet_view->set_zoom_scale( 125 ).  " 125% zoom

" View type controls the ruler/page-break overlay shown in Excel
lo_sheet_view->set_view( zcl_excel_sheet_view=>c_view_normal ).
" Other options: c_view_page_break_preview, c_view_page_layout

" Toggle display elements
lo_sheet_view->set_show_gridlines( abap_false ).        " Hide the cell grid
lo_sheet_view->set_show_row_col_headers( abap_false ).  " Hide row numbers / column letters
lo_sheet_view->set_show_zeros( abap_false ).            " Display zero values as blank
```

## Worksheet Data Organisation

### Named Ranges

```abap
" Create named ranges for easier formula reference
DATA: lo_range TYPE REF TO zcl_excel_range.

" Named ranges are registered at the workbook level, not the worksheet level
lo_range = lo_excel->add_new_range( ).
lo_range->set_name( 'SalesData' ).
lo_range->set_value( 'Summary!$A$1:$E$100' ).  " Use fully-qualified sheet reference

" Reference the named range by name in any formula
lo_worksheet->set_cell_formula(
  ip_column  = 'F'
  ip_row     = 1
  ip_formula = 'SUM(SalesData)'
).
```

### Data Validation

```abap
" Add a drop-down list validation to a column
DATA: lo_data_validation TYPE REF TO zcl_excel_data_validation.

lo_data_validation = lo_worksheet->add_new_data_validation( ).
lo_data_validation->set_range( 'B2:B100' ).                             " Apply to the whole column
lo_data_validation->set_type( zcl_excel_data_validation=>c_type_list ). " Drop-down list
lo_data_validation->set_formula1( 'North,South,East,West' ).            " Comma-separated options
lo_data_validation->set_allow_blank( abap_false ).                      " Mandatory field
lo_data_validation->set_show_dropdown( abap_true ).                     " Show the arrow button

" Validation error message shown when the user enters an invalid value
lo_data_validation->set_error_title( 'Invalid Region' ).
lo_data_validation->set_error( 'Please select a valid region from the dropdown.' ).
```

## Performance Considerations

### Efficient Worksheet Operations

```abap
" Batch operations for better performance
METHOD populate_worksheet_efficiently.
  " Rule 1: Minimise worksheet switches — complete all work on one sheet
  "         before moving to the next.

  " Rule 2: Use bind_table for large datasets instead of cell-by-cell loops
  lo_worksheet->bind_table( ip_table = lt_large_data ).

  " Rule 3: Create a style once and reuse it — do not call add_new_style inside a loop
  DATA(lo_header_style) = lo_excel->add_new_style( ).
  " ... configure the style here ...

  " Apply the pre-created style reference to every header cell
  LOOP AT lt_headers INTO DATA(ls_header).
    lo_worksheet->set_cell(
      ip_column = ls_header-column
      ip_row    = 1
      ip_value  = ls_header-text
      ip_style  = lo_header_style  " Reuse — no new style object per cell
    ).
  ENDLOOP.

  " Rule 4: Release references when done to free memory
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
- **[AutoFilter](/guide/autofilter)** - Add drop-down column filters
- **[Template Filling](/guide/template-filling)** - Fill pre-designed Excel templates
- **[Changelog](/guide/changelog)** - Full history of recent changes

## Common Worksheet Patterns

### Multi-Sheet Report Structure

```abap
" Standard pattern for a multi-sheet report with clear separation of concerns
METHOD create_multi_sheet_report.
  " Sheet 1: high-level KPIs and charts for management
  DATA(lo_summary) = lo_excel->add_new_worksheet( ).
  lo_summary->set_title( 'Executive Summary' ).

  " Sheet 2: full row-level data for analysts
  DATA(lo_details) = lo_excel->add_new_worksheet( ).
  lo_details->set_title( 'Detailed Data' ).

  " Sheet 3: visualisations derived from the detail data
  DATA(lo_charts) = lo_excel->add_new_worksheet( ).
  lo_charts->set_title( 'Analysis' ).

  " Populate each sheet via dedicated helper methods to keep the main method clean
  setup_summary_sheet( lo_summary ).
  setup_details_sheet( lo_details ).
  setup_charts_sheet( lo_charts ).
ENDMETHOD.
```

This guide covers the essential techniques for managing worksheets in abap2xlsx. Proper worksheet organisation is key to creating professional, navigable Excel reports.
