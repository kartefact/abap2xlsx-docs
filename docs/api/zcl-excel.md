# zcl_excel API Reference

Complete API reference for the main Excel workbook class in abap2xlsx.

## Overview

The `zcl_excel` class is the central class in abap2xlsx that represents an Excel workbook <cite>src/zcl_excel.clas.xml:6-8</cite>. It provides methods for managing worksheets, styles, drawings, ranges, and other workbook-level features.

## Class Definition

```abap
CLASS zcl_excel DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC
```

The class implements several interfaces for workbook properties, protection, and VBA project support <cite>src/zcl_excel.clas.abap:91-210</cite>.

## Constructor

### CREATE OBJECT

```abap
DATA: lo_excel TYPE REF TO zcl_excel.
CREATE OBJECT lo_excel.
```

The constructor initializes the workbook with default properties and creates the initial worksheet collection.

## Worksheet Management

### add_new_worksheet( )

Creates and adds a new worksheet to the workbook.

**Parameters:**

- `ip_worksheet_name` (optional) - Name for the new worksheet

**Returns:** Reference to `zcl_excel_worksheet`

```abap
DATA(lo_worksheet) = lo_excel->add_new_worksheet( ip_worksheet_name = 'Sales Data' ).
```

### get_worksheet_by_name( )

Retrieves a worksheet by its name <cite>src/zcl_excel.clas.abap:136-140</cite>.

**Parameters:**

- `ip_sheet_name` - Name of the worksheet to retrieve

**Returns:** Reference to `zcl_excel_worksheet`

```abap
DATA(lo_worksheet) = lo_excel->get_worksheet_by_name( 'Sales Data' ).
```

### get_worksheet_by_index( )

Retrieves a worksheet by its index position <cite>src/zcl_excel.clas.abap:131-135</cite>.

**Parameters:**

- `iv_index` - Index position (1-based)

**Returns:** Reference to `zcl_excel_worksheet`

```abap
DATA(lo_worksheet) = lo_excel->get_worksheet_by_index( 1 ).
```

### delete_worksheet( )

Removes a worksheet from the workbook.

**Parameters:**

- `io_worksheet` - Reference to worksheet to delete

**Raises:** `zcx_excel` if attempting to delete the last remaining worksheet

### get_worksheets_size( )

Returns the number of worksheets in the workbook <cite>src/zcl_excel.clas.abap:128-130</cite>.

**Returns:** Integer count of worksheets

## Active Worksheet Management

### set_active_sheet_index( )

Sets the active worksheet by index <cite>src/zcl_excel.clas.abap:141-145</cite>.

**Parameters:**

- `i_active_worksheet` - Index of worksheet to make active

**Raises:** `zcx_excel` if index is invalid

### set_active_sheet_index_by_name( )

Sets the active worksheet by name <cite>src/zcl_excel.clas.abap:146-148</cite>.

**Parameters:**

- `i_worksheet_name` - Name of worksheet to make active

### get_active_worksheet( )

Returns reference to the currently active worksheet.

**Returns:** Reference to `zcl_excel_worksheet`

## Style Management

### add_new_style( )

Creates a new style object for the workbook.

**Returns:** Reference to `zcl_excel_style`

```abap
DATA(lo_style) = lo_excel->add_new_style( ).
lo_style->font->bold = abap_true.
lo_style->font->color->set_rgb( 'FF0000' ).
```

### get_style_from_guid( )

Retrieves a style by its GUID <cite>src/zcl_excel.clas.abap:100-104</cite>.

**Parameters:**

- `ip_guid` - Style GUID

**Returns:** Reference to `zcl_excel_style`

### get_styles_iterator( )

Returns an iterator for all styles in the workbook <cite>src/zcl_excel.clas.abap:97-99</cite>.

**Returns:** Reference to `zcl_excel_collection_iterator`

### set_default_style( )

Sets the default style for the workbook <cite>src/zcl_excel.clas.abap:149-153</cite>.

**Parameters:**

- `ip_style` - Style GUID to use as default

**Raises:** `zcx_excel` if style GUID is invalid

## Drawing and Media Management

### add_new_drawing( )

Creates a new drawing object <cite>src/zcl_excel.clas.abap:205-210</cite>.

**Parameters:**

- `ip_type` (optional) - Drawing type
- `ip_title` (optional) - Drawing title

**Returns:** Reference to `zcl_excel_drawing`

```abap
DATA(lo_drawing) = lo_excel->add_new_drawing( 
  ip_type = zcl_excel_drawing=>type_image 
  ip_title = 'Company Logo'
).
```

## Range Management

### add_new_range( )

Creates a new named range in the workbook.

**Returns:** Reference to `zcl_excel_range`

```abap
DATA(lo_range) = lo_excel->add_new_range( ).
lo_range->set_name( 'SalesData' ).
lo_range->set_value( 'Sheet1!A1:D100' ).
```

## Table Management

### add_new_table( )

Creates a new table object for the workbook.

**Returns:** Reference to `zcl_excel_table`

## Autofilter Management

### add_new_autofilter( )

Creates a new autofilter for a worksheet <cite>src/zcl_excel.clas.abap:192-195</cite>.

**Parameters:**

- `io_sheet` - Worksheet to add autofilter to

**Returns:** Reference to `zcl_excel_autofilter`

## Comment Management

### add_new_comment( )

Creates a new comment object <cite>src/zcl_excel.clas.abap:198-202</cite>.

**Returns:** Reference to `zcl_excel_comment`

## Document Properties

The class implements `zif_excel_book_properties` interface for document metadata <cite>src/zcl_excel.clas.abap:662-674</cite>:

### Properties Available

- `application` - Application name
- `appversion` - Application version
- `created` - Creation timestamp
- `creator` - Document creator
- `description` - Document description
- `modified` - Last modified timestamp
- `lastmodifiedby` - Last modified by user

```abap
lo_excel->zif_excel_book_properties~creator = 'John Doe'.
lo_excel->zif_excel_book_properties~title = 'Sales Report 2024'.
lo_excel->zif_excel_book_properties~description = 'Monthly sales analysis'.
```

## Workbook Protection

The class implements `zif_excel_book_protection` interface for security <cite>src/zcl_excel.clas.abap:677-684</cite>:

### Protection Methods

- `set_protection_structure()` - Protect workbook structure
- `set_protection_windows()` - Protect workbook windows
- `set_workbook_password()` - Set workbook password

## Theme Support

### set_theme( )

Sets the theme for the workbook <cite>src/zcl_excel.clas.abap:154-156</cite>.

**Parameters:**

- `io_theme` - Reference to `zcl_excel_theme`

### get_theme( )

Retrieves the current theme <cite>src/zcl_excel.clas.abap:119-121</cite>.

**Returns:** Reference to `zcl_excel_theme`

## Template Support

### fill_template( )

Fills the workbook using template data <cite>src/zcl_excel.clas.abap:157-161</cite>.

**Parameters:**

- `iv_data` - Reference to `zcl_excel_template_data`

**Raises:** `zcx_excel` if template processing fails

## VBA Project Support

The class implements `zif_excel_book_vba_project` interface for macro support <cite>src/zcl_excel.clas.abap:687-699</cite>:

### VBA Methods

- `set_codename()` - Set VBA project codename
- `set_codename_pr()` - Set VBA project codename prefix
- `set_vbaproject()` - Set VBA project content

## Version Information

### version (Static Constant)

Returns the current library version.

```abap
DATA(lv_version) = zcl_excel=>version.
MESSAGE |abap2xlsx version: { lv_version }| TYPE 'I'.
```

## Usage Examples

### Complete Workbook Creation

```abap
" Create workbook with multiple features
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_style TYPE REF TO zcl_excel_style.

CREATE OBJECT lo_excel.

" Set document properties
lo_excel->zif_excel_book_properties~creator = sy-uname.
lo_excel->zif_excel_book_properties~title = 'Sales Analysis'.
lo_excel->zif_excel_book_properties~description = 'Q4 2024 Sales Data'.

" Create worksheet
lo_worksheet = lo_excel->add_new_worksheet( 'Q4 Sales' ).

" Create and apply styles
lo_style = lo_excel->add_new_style( ).
lo_style->font->bold = abap_true.
lo_style->font->size = 14.

" Add data and formatting
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Sales Report' ).
lo_worksheet->set_cell_style( ip_column = 'A' ip_row = 1 ip_style = lo_style ).

" Set as active worksheet
lo_excel->set_active_sheet_index( 1 ).
```

## Error Handling

Most methods that can fail raise `zcx_excel` exceptions. Always wrap critical operations in TRY-CATCH blocks:

```abap
TRY.
    lo_excel->set_active_sheet_index( 5 ).
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE lx_excel->get_text( ) TYPE 'E'.
ENDTRY.
```

## Best Practices

1. **Resource Management**: Clear object references when done
2. **Error Handling**: Use TRY-CATCH for operations that can fail
3. **Performance**: Reuse style objects instead of creating duplicates
4. **Memory**: Use appropriate writers for large datasets

## Related Classes

- `zcl_excel_worksheet` - Individual worksheet management
- `zcl_excel_style` - Cell and range formatting
- `zcl_excel_writer_2007` - Excel file generation
- `zcl_excel_reader_2007` - Excel file reading
