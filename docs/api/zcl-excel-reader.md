# zcl_excel_reader API Reference

Complete API reference for the Excel reader classes in abap2xlsx.

## Overview

The reader classes in abap2xlsx are responsible for parsing Excel files and converting them into the in-memory Excel object model. The library provides reader implementations for different Excel formats and use cases <cite>src/zcl_excel_reader_2007.clas.abap:91-150</cite>.

## Reader Interface

### zif_excel_reader

All reader classes implement the `zif_excel_reader` interface, which provides a consistent API for file reading.

**Core Methods:**

- `load_file()` - Loads Excel file from various sources
- `load()` - Loads from xstring data

## Standard Excel Reader (XLSX)

### zcl_excel_reader_2007

The primary reader class for reading Excel 2007+ (.xlsx) files <cite>src/zcl_excel_reader_2007.clas.abap:271-390</cite>.

#### Constructor

```abap
DATA: lo_reader TYPE REF TO zcl_excel_reader_2007.
CREATE OBJECT lo_reader.
```

#### load_file( )

Loads an Excel file from various sources and returns a workbook object.

**Parameters:**

- `i_filename` (optional) - File path for application server or local file
- `i_xlsx_binary` (optional) - Excel file as xstring
- `i_use_alternate_zip` (optional) - Alternative ZIP class for processing

**Returns:** Reference to `zcl_excel` workbook object

**Raises:** `zcx_excel`

```abap
" Load from application server
DATA(lo_excel) = lo_reader->load_file( i_filename = '/path/to/file.xlsx' ).

" Load from xstring
DATA(lo_excel) = lo_reader->load_file( i_xlsx_binary = lv_file_data ).
```

#### Internal Architecture

The reader follows a structured approach to Excel file parsing <cite>src/zcl_excel_reader_2007.clas.abap:1831-1890</cite>:

1. **ZIP Archive Creation**: Extracts the Excel file structure
2. **Relationship Processing**: Parses document relationships
3. **Shared Strings Loading**: Processes shared string table
4. **Styles Processing**: Loads formatting definitions
5. **Worksheet Parsing**: Processes individual sheet data
6. **Formula Resolution**: Resolves shared and array formulas
7. **Drawing Integration**: Handles images and charts
8. **Theme Loading**: Processes Excel themes

#### Key Internal Methods

##### load_shared_strings( )

Processes the shared strings table for text optimization <cite>src/zcl_excel_reader_2007.clas.abap:147-150</cite>.

**Parameters:**

- `ip_path` - Path to shared strings XML

**Raises:** `zcx_excel`

##### load_worksheet_tables( )

Loads Excel table definitions from worksheets <cite>src/zcl_excel_reader_2007.clas.abap:278-286</cite>.

**Parameters:**

- `io_ixml_worksheet` - Worksheet XML document
- `io_worksheet` - Target worksheet object
- `iv_dirname` - Directory path
- `it_tables` - Table relationship information

**Raises:** `zcx_excel`

##### load_theme( )

Processes Excel theme information <cite>src/zcl_excel_reader_2007.clas.abap:304-309</cite>.

**Parameters:**

- `iv_path` - Path to theme XML
- `ip_excel` - Target workbook object

**Raises:** `zcx_excel`

## Data Type Handling

### Shared String Processing

The reader maintains a shared strings table for efficient text storage <cite>src/zcl_excel_reader_2007.clas.abap:103-108</cite>:

```abap
TYPES:
  BEGIN OF t_shared_string,
    value TYPE string,
    rtf   TYPE zexcel_t_rtf,
  END OF t_shared_string .
TYPES:
  t_shared_strings TYPE STANDARD TABLE OF t_shared_string WITH DEFAULT KEY .
```

### Formula Reference Management

The reader tracks formula references for proper resolution <cite>src/zcl_excel_reader_2007.clas.abap:91-101</cite>:

```abap
TYPES:
  BEGIN OF ty_ref_formulae,
    sheet   TYPE REF TO zcl_excel_worksheet,
    row     TYPE i,
    column  TYPE i,
    si      TYPE i,
    ref     TYPE string,
    formula TYPE string,
  END   OF ty_ref_formulae .
```

## Column Processing

### Column Attribute Handling

The reader processes column definitions with comprehensive attribute support <cite>src/zcl_excel_reader_2007.clas.abap:2731-2790</cite>:

- **Width Settings**: Custom width and auto-sizing
- **Visibility**: Hidden and collapsed states
- **Outline Levels**: Grouping and hierarchy
- **Style Application**: Column-level formatting

## ZIP Archive Management

### Archive Processing

The reader uses a local ZIP archive class for file extraction <cite>src/zcl_excel_reader_2007.clas.abap:341-351</cite>:

```abap
METHODS create_zip_archive
  IMPORTING
    !i_xlsx_binary       TYPE xstring
    !i_use_alternate_zip TYPE seoclsname OPTIONAL
  RETURNING
    VALUE(e_zip)         TYPE REF TO lcl_zip_archive
  RAISING
    zcx_excel .
```

### File Source Handling

The reader supports multiple file sources <cite>src/zcl_excel_reader_2007.clas.abap:352-365</cite>:

- **Application Server**: Direct file system access
- **Local Files**: Frontend file upload
- **Binary Data**: In-memory xstring processing

## Namespace Support

### XML Namespace Handling

The reader defines comprehensive namespace support for Excel XML processing <cite>src/zcl_excel_reader_2007.clas.abap:316-337</cite>:

```abap
CONSTANTS: BEGIN OF namespace,
             main             TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
             relationships    TYPE string VALUE 'http://schemas.openxmlformats.org/package/2006/relationships',
             drawing          TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
             worksheet        TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
           END OF namespace.
```

## Usage Examples

### Basic File Reading

```abap
" Read Excel file from application server
DATA: lo_reader TYPE REF TO zcl_excel_reader_2007,
      lo_excel TYPE REF TO zcl_excel.

CREATE OBJECT lo_reader.

TRY.
    lo_excel = lo_reader->load_file( i_filename = '/tmp/report.xlsx' ).
    
    " Access worksheet data
    DATA(lo_worksheet) = lo_excel->get_active_worksheet( ).
    DATA(lv_cell_value) = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).
    
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE lx_excel->get_text( ) TYPE 'E'.
ENDTRY.
```

### Reading with Alternative ZIP Handler

```abap
" Use alternative ZIP processing for specific requirements
DATA(lo_excel) = lo_reader->load_file( 
  i_xlsx_binary = lv_file_data
  i_use_alternate_zip = 'ZCL_CUSTOM_ZIP_HANDLER'
).
```

### Processing Multiple Worksheets

```abap
" Iterate through all worksheets
DATA: lo_worksheets TYPE REF TO zcl_excel_worksheets,
      lo_iterator TYPE REF TO zcl_excel_collection_iterator.

lo_worksheets = lo_excel->get_worksheets( ).
lo_iterator = lo_worksheets->get_iterator( ).

WHILE lo_iterator->has_next( ) = abap_true.
  DATA(lo_worksheet) = CAST zcl_excel_worksheet( lo_iterator->get_next( ) ).
  
  " Process worksheet data
  process_worksheet_data( lo_worksheet ).
ENDWHILE.
```

## Error Handling

### Exception Management

All reader methods can raise `zcx_excel` exceptions. Implement comprehensive error handling:

```abap
TRY.
    DATA(lo_excel) = lo_reader->load_file( i_filename = lv_file_path ).
    
  CATCH zcx_excel INTO DATA(lx_excel).
    CASE lx_excel->textid.
      WHEN zcx_excel=>file_not_found.
        MESSAGE 'Excel file not found' TYPE 'E'.
      WHEN zcx_excel=>invalid_file_format.
        MESSAGE 'Invalid Excel file format' TYPE 'E'.
      WHEN OTHERS.
        MESSAGE lx_excel->get_text( ) TYPE 'E'.
    ENDCASE.
ENDTRY.
```

## Performance Considerations

### Memory Management

- **Large Files**: Consider memory limitations when reading large Excel files
- **Streaming**: The reader loads entire file into memory - not suitable for very large files
- **Resource Cleanup**: Clear object references after processing

### Processing Optimization

- **Selective Reading**: Focus on specific worksheets or ranges when possible
- **Formula Resolution**: Complex formulas may impact reading performance
- **Image Processing**: Embedded images and charts add processing overhead

## Best Practices

1. **Error Handling**: Always wrap read operations in TRY-CATCH blocks
2. **File Validation**: Verify file format before processing
3. **Memory**: Monitor memory usage with large files
4. **Performance**: Test with representative file sizes
5. **Compatibility**: Ensure Excel files are in supported format (2007+)

## Limitations

### Current Limitations

- **Excel 97-2003**: Does not support .xls format (only .xlsx)
- **Macro Content**: Limited VBA macro support (preservation only)
- **Advanced Features**: Some advanced Excel features may not be fully supported
- **File Size**: Memory constraints for very large files

### Workarounds

- **Format Conversion**: Convert .xls files to .xlsx before processing
- **File Splitting**: Break large files into smaller chunks
- **Selective Processing**: Read only required worksheets or ranges

## Related Classes

- `zcl_excel` - Target workbook object
- `zcl_excel_worksheet` - Individual worksheet data
- `zcl_excel_writer_2007` - Corresponding writer functionality
- `zcl_excel_common` - Utility functions for data conversion
