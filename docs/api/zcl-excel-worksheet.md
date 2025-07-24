# zcl_excel_worksheet API Reference

Complete API reference for the Excel worksheet class in abap2xlsx.

## Overview

The `zcl_excel_worksheet` class represents an individual Excel worksheet within a workbook <cite>src/zcl_excel_worksheet.clas.abap:1-5</cite>. It provides comprehensive methods for managing cell data, styling, formulas, and worksheet-specific features like merged cells, data validation, and page setup.

## Class Definition

```abap
CLASS zcl_excel_worksheet DEFINITION
  PUBLIC
  CREATE PUBLIC
```

The class implements several interfaces for sheet properties, protection, print settings, and VBA project support <cite>src/zcl_excel_worksheet.clas.abap:13-16</cite>.

## Constructor

### constructor( )

Initializes a new worksheet with default settings <cite>src/zcl_excel_worksheet.clas.abap:2101-2118</cite>.

**Parameters:**

- `ip_excel` - Reference to parent workbook
- `ip_worksheet_name` (optional) - Name for the worksheet

```abap
DATA(lo_worksheet) = NEW zcl_excel_worksheet( ip_excel = lo_excel ).
```

## Cell Operations

### set_cell( )

Sets cell value, formula, style, and other properties <cite>src/zcl_excel_worksheet.clas.abap:497-513</cite>.

**Parameters:**

- `ip_columnrow` (optional) - Excel notation (e.g., 'A1')
- `ip_column` (optional) - Column number or letter
- `ip_row` (optional) - Row number
- `ip_value` (optional) - Cell value
- `ip_formula` (optional) - Excel formula
- `ip_style` (optional) - Style reference or GUID
- `ip_hyperlink` (optional) - Hyperlink object
- `ip_data_type` (optional) - Data type override
- `ip_abap_type` (optional) - ABAP type information
- `ip_currency` (optional) - Currency code
- `it_rtf` (optional) - Rich text formatting
- `ip_column_formula_id` (optional) - Column formula ID
- `ip_conv_exit_length` (optional) - Conversion exit handling

**Raises:** `zcx_excel`

```abap
" Set simple value
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello World' ).

" Set with formula and style
lo_worksheet->set_cell( 
  ip_columnrow = 'B2'
  ip_formula = 'SUM(A1:A10)'
  ip_style = lo_style
).
```

### get_cell( )

Retrieves cell value and properties <cite>src/zcl_excel_worksheet.clas.abap:340-353</cite>.

**Parameters:**

- `ip_columnrow` (optional) - Excel notation
- `ip_column` (optional) - Column identifier
- `ip_row` (optional) - Row number

**Exports:**

- `ep_value` - Cell value
- `ep_rc` - Return code
- `ep_style` - Style object reference
- `ep_guid` - Style GUID
- `ep_formula` - Cell formula
- `et_rtf` - Rich text formatting

**Raises:** `zcx_excel`

### set_cell_formula( )

Sets only the formula for a cell <cite>src/zcl_excel_worksheet.clas.abap:514-521</cite>.

**Parameters:**

- `ip_columnrow` (optional) - Excel notation
- `ip_column` (optional) - Column identifier
- `ip_row` (optional) - Row number
- `ip_formula` - Excel formula

**Raises:** `zcx_excel`

### set_cell_style( )

Sets only the style for a cell <cite>src/zcl_excel_worksheet.clas.abap:522-529</cite>.

**Parameters:**

- `ip_columnrow` (optional) - Excel notation
- `ip_column` (optional) - Column identifier
- `ip_row` (optional) - Row number
- `ip_style` - Style reference or GUID

**Raises:** `zcx_excel`

## Range Operations

### set_area( )

Sets values, formulas, and styles for a range of cells <cite>src/zcl_excel_worksheet.clas.abap:653-669</cite>.

**Parameters:**

- `ip_range` (optional) - Range notation (e.g., 'A1:C10')
- `ip_column_start` (optional) - Start column
- `ip_column_end` (optional) - End column
- `ip_row` (optional) - Start row
- `ip_row_to` (optional) - End row
- `ip_value` (optional) - Value to set
- `ip_formula` (optional) - Formula to set
- `ip_style` (optional) - Style to apply
- `ip_hyperlink` (optional) - Hyperlink object
- `ip_data_type` (optional) - Data type
- `ip_abap_type` (optional) - ABAP type
- `ip_merge` (optional) - Merge cells flag
- `ip_area` (optional) - Area type

**Raises:** `zcx_excel`

### set_area_formula( )

Sets formula for a range of cells <cite>src/zcl_excel_worksheet.clas.abap:630-641</cite>.

### set_area_style( )

Sets style for a range of cells <cite>src/zcl_excel_worksheet.clas.abap:642-652</cite>.

## Data Binding

### bind_table( )

Binds an internal table to the worksheet <cite>src/zcl_excel_worksheet.clas.abap:170-182</cite>.

**Parameters:**

- `ip_table` - Internal table to bind
- `it_field_catalog` (optional) - Field catalog for customization
- `is_table_settings` (optional) - Table positioning and styling

**Raises:** `zcx_excel`

```abap
lo_worksheet->bind_table(
  ip_table = lt_sales_data
  is_table_settings = VALUE #(
    top_left_column = 'A'
    top_left_row = 2
    table_style = zcl_excel_table=>builtinstyle_medium9
  )
).
```

### bind_alv( )

Binds ALV grid data to the worksheet <cite>src/zcl_excel_worksheet.clas.abap:138-147</cite>.

**Parameters:**

- `io_alv` - ALV grid object reference
- `it_table` - Table data
- `i_top` - Starting row (default 1)
- `i_left` - Starting column (default 1)
- `table_style` (optional) - Table style
- `i_table` - Create as Excel table (default true)

**Raises:** `zcx_excel`

### convert_to_table( )

Converts worksheet data back to internal table <cite>src/zcl_excel_worksheet.clas.abap:2121-2160</cite>.

**Parameters:**

- `it_field_catalog` (optional) - Field catalog for conversion
- `iv_begin_row` - Starting row (default 2)
- `iv_end_row` - Ending row (default 0 = all)

**Exports:**

- `et_data` - Converted internal table
- `er_data` - Reference to string-based table

**Raises:** `zcx_excel`

## Column and Row Management

### add_new_column( )

Creates a new column definition <cite>src/zcl_excel_worksheet.clas.abap:115-121</cite>.

**Parameters:**

- `ip_column` - Column identifier

**Returns:** Reference to `zcl_excel_column`

**Raises:** `zcx_excel`

### get_column( )

Retrieves column object <cite>src/zcl_excel_worksheet.clas.abap:354-360</cite>.

**Parameters:**

- `ip_column` - Column identifier

**Returns:** Reference to `zcl_excel_column`

**Raises:** `zcx_excel`

### set_column_width( )

Sets column width <cite>src/zcl_excel_worksheet.clas.abap:530-536</cite>.

**Parameters:**

- `ip_column` - Column identifier
- `ip_width_fix` - Fixed width (default 0)
- `ip_width_autosize` - Auto-size flag (default 'X')

**Raises:** `zcx_excel`

### add_new_row( )

Creates a new row definition <cite>src/zcl_excel_worksheet.clas.abap:133-137</cite>.

**Parameters:**

- `ip_row` - Row number

**Returns:** Reference to `zcl_excel_row`

### set_row_height( )

Sets row height <cite>src/zcl_excel_worksheet.clas.abap:565-570</cite>.

**Parameters:**

- `ip_row` - Row number
- `ip_height_fix` - Fixed height

**Raises:** `zcx_excel`

## Merged Cells

### set_merge( )

Creates merged cell range <cite>src/zcl_excel_worksheet.clas.abap:545-556</cite>.

**Parameters:**

- `ip_range` (optional) - Range notation
- `ip_column_start` (optional) - Start column
- `ip_column_end` (optional) - End column
- `ip_row` (optional) - Start row
- `ip_row_to` (optional) - End row
- `ip_style` (optional) - Style to apply
- `ip_value` (optional) - Value for merged cell
- `ip_formula` (optional) - Formula for merged cell

**Raises:** `zcx_excel`

### is_cell_merged( )

Checks if a cell is part of a merged range <cite>src/zcl_excel_worksheet.clas.abap:489-496</cite>.

**Parameters:**

- `ip_column` - Column identifier
- `ip_row` - Row number

**Returns:** Boolean indicating if cell is merged

**Raises:** `zcx_excel`

### get_merge( )

Returns all merged cell ranges <cite>src/zcl_excel_worksheet.clas.abap:443-447</cite>.

**Returns:** String table of merge ranges

**Raises:** `zcx_excel`

## Data Validation

### add_new_data_validation( )

Creates a new data validation rule <cite>src/zcl_excel_worksheet.clas.abap:127-129</cite>.

**Returns:** Reference to `zcl_excel_data_validation`

### get_data_validations_iterator( )

Returns iterator for data validation rules <cite>src/zcl_excel_worksheet.clas.abap:374-376</cite>.

**Returns:** Reference to `zcl_excel_collection_iterator`

### get_data_validations_size( )

Returns count of data validation rules <cite>src/zcl_excel_worksheet.clas.abap:377-379</cite>.

**Returns:** Integer count

## Conditional Formatting

### add_new_style_cond( )

Creates a new conditional formatting rule <cite>src/zcl_excel_worksheet.clas.abap:122-126</cite>.

**Parameters:**

- `ip_dimension_range` - Target range (default 'A1')

**Returns:** Reference to `zcl_excel_style_cond`

### get_style_cond( )

Retrieves conditional formatting rule by GUID <cite>src/zcl_excel_worksheet.clas.abap:470-474</cite>.

**Parameters:**

- `ip_guid` - Style condition GUID

**Returns:** Reference to `zcl_excel_style_cond`

### get_style_cond_iterator( )

Returns iterator for conditional formatting rules <cite>src/zcl_excel_worksheet.clas.abap:371-373</cite>.

**Returns:** Reference to `zcl_excel_collection_iterator`

## Worksheet Properties

### set_title( )

Sets worksheet name <cite>src/zcl_excel_worksheet.clas.abap:604-608</cite>.

**Parameters:**

- `ip_title` - Worksheet title

**Raises:** `zcx_excel`

### get_title( )

Gets worksheet name <cite>src/zcl_excel_worksheet.clas.abap:484-488</cite>.

**Parameters:**

- `ip_escaped` - Return escaped title (default false)

**Returns:** Worksheet title

### set_tabcolor( )

Sets tab color <cite>src/zcl_excel_worksheet.clas.abap:589-591</cite>.

**Parameters:**

- `iv_tabcolor` - Tab color structure

### get_tabcolor( )

Gets tab color <cite>src/zcl_excel_worksheet.clas.abap:475-477</cite>.

**Returns:** Tab color structure

## Page Setup and Print Settings

### freeze_panes( )

Freezes rows and/or columns <cite>src/zcl_excel_worksheet.clas.abap:2463-2479</cite>.

**Parameters:**

- `ip_num_columns` (optional) - Number of columns to freeze
- `ip_num_rows` (optional) - Number of rows to freeze

**Raises:** `zcx_excel`

### get_freeze_cell( )

Gets freeze pane position <cite>src/zcl_excel_worksheet.clas.abap:417-420</cite>.

**Exports:**

- `ep_row` - Freeze row position
- `ep_column` - Freeze column position

### set_print_gridlines( )

Sets print gridlines option <cite>src/zcl_excel_worksheet.clas.abap:562-564</cite>.

**Parameters:**

- `i_print_gridlines` - Print gridlines flag

### set_show_gridlines( )

Sets show gridlines option <cite>src/zcl_excel_worksheet.clas.abap:583-585</cite>.

**Parameters:**

- `i_show_gridlines` - Show gridlines flag

## Dimension and Navigation

### get_dimension_range( )

Gets the used range of the worksheet <cite>src/zcl_excel_worksheet.clas.abap:394-398</cite>.

**Returns:** Range string (e.g., 'A1:D10')

**Raises:** `zcx_excel`

### get_highest_row( )

Gets the highest used row number <cite>src/zcl_excel_worksheet.clas.abap:429-433</cite>.

**Returns:** Row number

**Raises:** `zcx_excel`

### get_highest_column( )

Gets the highest used column number <cite>src/zcl_excel_worksheet.clas.abap:424-428</cite>.

**Returns:** Column number

**Raises:** `zcx_excel`

## Error Handling

### set_ignored_errors( )

Sets error types to ignore <cite>src/zcl_excel_worksheet.clas.abap:542-544</cite>.

**Parameters:**

- `it_ignored_errors` - Table of ignored error configurations

### get_ignored_errors( )

Gets ignored error configurations <cite>src/zcl_excel_worksheet.clas.abap:440-442</cite>.

**Returns:** Table of ignored error configurations

I'll continue from where the `zcl_excel_worksheet` API documentation left off. Here's the completion of that file:

## `docs/api/zcl-excel-worksheet.md` (continued)

```markdown
## Usage Examples

### Complete Worksheet Setup

```abap
" Create and configure worksheet
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_style TYPE REF TO zcl_excel_style.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( 'Sales Report' ).

" Set worksheet properties
lo_worksheet->set_tabcolor( VALUE #( rgb = 'FF0000' ) ).
lo_worksheet->freeze_panes( ip_num_rows = 1 ip_num_columns = 2 ).

" Add data with styling
lo_style = lo_excel->add_new_style( ).
lo_style->font->bold = abap_true.

lo_worksheet->set_cell( 
  ip_column = 'A' 
  ip_row = 1 
  ip_value = 'Product' 
  ip_style = lo_style 
).

" Bind table data
lo_worksheet->bind_table(
  ip_table = lt_sales_data
  is_table_settings = VALUE #(
    top_left_column = 'A'
    top_left_row = 2
  )
).

" Add data validation
DATA(lo_validation) = lo_worksheet->add_new_data_validation( ).
lo_validation->type = zcl_excel_data_validation=>c_type_list.
lo_validation->formula1 = '"High,Medium,Low"'.

" Add conditional formatting
DATA(lo_cond_format) = lo_worksheet->add_new_style_cond( 'D2:D100' ).
lo_cond_format->set_rule_type( zcl_excel_style_cond=>c_rule_top10 ).
```

### Working with Merged Cells

```abap
" Create merged header
lo_worksheet->set_merge(
  ip_range = 'A1:D1'
  ip_value = 'Quarterly Sales Report'
  ip_style = lo_header_style
).

" Check if cell is merged
IF lo_worksheet->is_cell_merged( ip_column = 'B' ip_row = 1 ) = abap_true.
  MESSAGE 'Cell B1 is part of a merged range' TYPE 'I'.
ENDIF.
```

### Data Conversion

```abap
" Convert worksheet back to internal table
lo_worksheet->convert_to_table(
  IMPORTING
    et_data = lt_converted_data
    er_data = lr_string_data
).
```

## Error Handling

Most worksheet methods can raise `zcx_excel` exceptions. Always use proper error handling:

```abap
TRY.
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Test' ).
    lo_worksheet->freeze_panes( ip_num_rows = 5 ).
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE lx_excel->get_text( ) TYPE 'E'.
ENDTRY.
```

## Best Practices

1. **Performance**: Use `bind_table()` for large datasets instead of individual `set_cell()` calls
2. **Memory**: Clear large internal tables after binding to free memory
3. **Validation**: Always validate column/row parameters before use
4. **Styling**: Reuse style objects to avoid creating duplicates
5. **Error Handling**: Wrap worksheet operations in TRY-CATCH blocks

## Integration Points

The worksheet class integrates with several other abap2xlsx components:

- **Tables**: <cite>src/zcl_excel_worksheet.clas.abap:1324-1347</cite> - Automatic detection of table headers for styling
- **Data Validation**: <cite>src/zcl_excel_worksheet.clas.abap:894-898</cite> - Seamless integration with validation collection
- **Conditional Formatting**: <cite>src/zcl_excel_worksheet.clas.abap:916-919</cite> - Direct access to conditional formatting rules
- **ALV Conversion**: <cite>src/zcl_excel_worksheet.clas.abap:922-942</cite> - Built-in ALV grid conversion support

## Related Classes

- `zcl_excel` - Parent workbook class
- `zcl_excel_style` - Cell and range formatting
- `zcl_excel_data_validation` - Input validation rules
- `zcl_excel_style_cond` - Conditional formatting
- `zcl_excel_table` - Excel table functionality
- `zcl_excel_column` - Column-specific operations
- `zcl_excel_row` - Row-specific operations
