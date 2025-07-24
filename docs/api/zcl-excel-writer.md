# zcl_excel_writer API Reference

Complete API reference for the Excel writer classes in abap2xlsx.

## Overview

The writer classes in abap2xlsx are responsible for converting the in-memory Excel object model into actual Excel files. The library provides several writer implementations for different formats and use cases <cite>src/zcl_excel_writer_2007.clas.abap:1-10</cite>.

## Writer Interface

### zif_excel_writer

All writer classes implement the `zif_excel_writer` interface <cite>src/zcl_excel_writer_2007.clas.abap:9</cite>, which provides a consistent API for file generation.

**Core Method:**

- `write_file()` - Generates Excel file content

## Standard Excel Writer (XLSX)

### zcl_excel_writer_2007

The primary writer class for generating Excel 2007+ (.xlsx) files <cite>src/zcl_excel_writer_2007.clas.abap:1-4</cite>.

#### Constructor

```abap
DATA: lo_writer TYPE REF TO zcl_excel_writer_2007.
CREATE OBJECT lo_writer.
```

#### write_file( )

Generates the complete Excel file as an xstring.

**Parameters:**

- `io_excel` - Reference to `zcl_excel` workbook object

**Returns:** `xstring` containing the Excel file data

**Raises:** `zcx_excel`

```abap
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_writer TYPE REF TO zcl_excel_writer_2007,
      lv_file_data TYPE xstring.

CREATE OBJECT lo_writer.
lv_file_data = lo_writer->write_file( lo_excel ).
```

#### Internal Architecture

The writer follows a structured approach to Excel file generation <cite>src/zcl_excel_writer_2007.clas.abap:331-420</cite>:

1. **Archive Creation**: Creates ZIP container structure
2. **Content Types**: Defines MIME types for Excel components <cite>src/zcl_excel_writer_2007.clas.abap:597-630</cite>
3. **Relationships**: Establishes document relationships
4. **Document Properties**: Adds metadata and properties <cite>src/zcl_excel_writer_2007.clas.abap:931-990</cite>
5. **Workbook Structure**: Creates workbook.xml
6. **Styles**: Generates styles.xml with formatting definitions
7. **Shared Strings**: Optimizes text storage
8. **Worksheets**: Processes individual sheet data
9. **Media Content**: Handles images and charts <cite>src/zcl_excel_writer_2007.clas.abap:511-630</cite>

#### Key Internal Methods

##### create_xl_sheet( )

Creates individual worksheet XML content <cite>src/zcl_excel_writer_2007.clas.abap:125-132</cite>.

**Parameters:**

- `io_worksheet` - Worksheet object reference
- `iv_active` - Flag indicating if this is the active sheet

**Returns:** `xstring` containing worksheet XML

##### create_xl_styles( )

Generates the styles.xml file containing all formatting definitions <cite>src/zcl_excel_writer_2007.clas.abap:154-156</cite>.

**Returns:** `xstring` containing styles XML

##### create_xl_sharedstrings( )

Creates the shared strings XML for text optimization <cite>src/zcl_excel_writer_2007.clas.abap:122-124</cite>.

**Returns:** `xstring` containing shared strings XML

## Macro-Enabled Writer (XLSM)

### zcl_excel_writer_xlsm

Specialized writer for macro-enabled Excel files <cite>src/zcl_excel_writer_xlsm.clas.abap:1-4</cite>.

#### Inheritance Structure

```abap
CLASS zcl_excel_writer_xlsm DEFINITION
  PUBLIC
  INHERITING FROM zcl_excel_writer_2007
```

The XLSM writer extends the standard writer with VBA project support <cite>src/zcl_excel_writer_xlsm.clas.abap:37-45</cite>.

#### VBA Project Handling

The XLSM writer adds VBA project content to the Excel file:

```abap
" Add vbaProject.bin to zip
io_zip->add( name    = me->c_xl_vbaproject
             content = me->excel->zif_excel_book_vba_project~vbaproject ).
```

#### Usage Example

```abap
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_writer TYPE REF TO zcl_excel_writer_xlsm.

" Create macro-enabled writer
CREATE OBJECT lo_writer.

" Generate XLSM file
DATA(lv_xlsm_data) = lo_writer->write_file( lo_excel ).
```

## High-Performance Writer

### zcl_excel_writer_huge_file

Optimized writer for processing very large datasets <cite>src/zcl_excel_writer_huge_file.clas.abap:1-6</cite>.

#### Key Features

- **Memory Efficiency**: Uses streaming approach for large files
- **Simple Transformation**: Leverages XSLT for XML generation <cite>src/zcl_excel_writer_huge_file.clas.abap:41-115</cite>
- **Reduced Feature Set**: Focuses on core functionality for performance

#### Cell Data Structure

```abap
TYPES:
  BEGIN OF ty_cell,
    name    TYPE c LENGTH 10, "AAA1234567"
    style   TYPE i,
    type    TYPE c LENGTH 9,
    formula TYPE string,
    value   TYPE string,
  END OF ty_cell .
```

#### Usage Pattern

```abap
DATA: lo_huge_writer TYPE REF TO zcl_excel_writer_huge_file.

CREATE OBJECT lo_huge_writer.

" Process large dataset efficiently
lo_huge_writer->get_cells( i_row = lv_row i_index = lv_index ).
DATA(lv_file) = lo_huge_writer->write_file( lo_excel ).
```

## CSV Writer

### zcl_excel_writer_csv

Simple writer for CSV format output.

#### Usage

```abap
DATA: lo_csv_writer TYPE REF TO zcl_excel_writer_csv.

CREATE OBJECT lo_csv_writer.
DATA(lv_csv_data) = lo_csv_writer->write_file( lo_excel ).
```

## Writer Selection Guidelines

### Choose the Right Writer

```abap
METHOD select_appropriate_writer.
  DATA: lo_writer TYPE REF TO zif_excel_writer.

  " Decision logic for writer selection
  IF iv_has_macros = abap_true.
    " Use XLSM writer for macro support
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_xlsm.
    
  ELSEIF iv_row_count > 100000.
    " Use huge file writer for large datasets
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_huge_file.
    
  ELSEIF iv_simple_data = abap_true.
    " Use CSV writer for simple tabular data
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_csv.
    
  ELSE.
    " Use standard writer for normal cases
    CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  ENDIF.

  " Generate file
  DATA(lv_file_data) = lo_writer->write_file( io_excel ).
ENDMETHOD.
```

## Advanced Writer Features

### Custom Content Extension

The writer architecture supports extension through the `add_further_data_to_zip()` method <cite>src/zcl_excel_writer_2007.clas.abap:55-57</cite>:

```abap
" Override in child classes to add custom content
METHODS add_further_data_to_zip
  IMPORTING
    !io_zip TYPE REF TO cl_abap_zip .
```

### XML Document Utilities

The writer provides utility methods for XML processing <cite>src/zcl_excel_writer_2007.clas.abap:224-231</cite>:

```abap
" Create XML document
METHODS create_xml_document
  RETURNING
    VALUE(ro_document) TYPE REF TO if_ixml_document.

" Render XML to xstring
METHODS render_xml_document
  IMPORTING
    io_document       TYPE REF TO if_ixml_document
  RETURNING
    VALUE(ep_content) TYPE xstring.
```

## Error Handling

All writer methods can raise `zcx_excel` exceptions. Implement proper error handling:

```abap
TRY.
    DATA(lv_file) = lo_writer->write_file( lo_excel ).
    
  CATCH zcx_excel INTO DATA(lx_excel).
    MESSAGE lx_excel->get_text( ) TYPE 'E'.
ENDTRY.
```

## Performance Considerations

### Memory Management

- **Standard Writer**: Suitable for typical business reports (< 50,000 rows)
- **Huge File Writer**: Optimized for large datasets (> 100,000 rows)
- **Streaming**: Consider processing data in batches for very large files

### File Size Optimization

- **Shared Strings**: Automatically optimizes repeated text values
- **Style Reuse**: Reuse style objects to minimize file size
- **Image Compression**: Optimize images before adding to workbook

## Best Practices

1. **Writer Selection**: Choose appropriate writer based on requirements
2. **Error Handling**: Always wrap write operations in TRY-CATCH
3. **Memory**: Clear large objects after file generation
4. **Testing**: Test with representative data volumes
5. **Format**: Use XLSX for rich formatting, CSV for simple data exchange

## Related Classes

- `zcl_excel` - Source workbook object
- `zcl_excel_worksheet` - Individual worksheet data
- `zcl_excel_reader_2007` - Corresponding reader functionality
- `cl_abap_zip` - ZIP archive management for XLSX format
