# Debugging abap2xlsx Applications

This guide covers debugging techniques and tools for troubleshooting abap2xlsx applications.

## Common Debugging Scenarios

### Excel File Generation Issues

When Excel files are not generated correctly:

```abap
" Enable debug mode in writer
DATA: lo_writer TYPE REF TO zcl_excel_writer_2007.
CREATE OBJECT lo_writer.

" Check if workbook has worksheets
IF lo_excel->get_worksheets_size( ) = 0.
  MESSAGE 'No worksheets found' TYPE 'E'.
ENDIF.

" Verify worksheet content
DATA(lo_worksheet) = lo_excel->get_active_worksheet( ).
DATA(lv_cell_count) = lo_worksheet->get_highest_row( ).
WRITE: / 'Highest row:', lv_cell_count.
```

### Memory Issues with Large Files

For large datasets, monitor memory consumption:

```abap
" Check memory before processing
CALL FUNCTION 'SYSTEM_MEMORY_INFO'
  IMPORTING
    memory_available = DATA(lv_memory_before).

" Process data in chunks
DATA: lv_chunk_size TYPE i VALUE 1000.
DO.
  " Process chunk
  " Check memory periodically
  IF sy-index MOD 10 = 0.
    CALL FUNCTION 'SYSTEM_MEMORY_INFO'
      IMPORTING
        memory_available = DATA(lv_memory_current).
    
    IF lv_memory_current < lv_memory_before / 2.
      MESSAGE 'Low memory warning' TYPE 'W'.
    ENDIF.
  ENDIF.
ENDDO.
```

## Debugging Tools

### Using ZDEMO_EXCEL_CHECKER

The main diagnostic tool for verifying installation:

1. Execute `ZDEMO_EXCEL_CHECKER`
2. Review all test results
3. Focus on failed tests for specific issues

### Custom Debug Reports

Create custom debug reports to isolate issues:

```abap
REPORT zdebug_excel_test.

DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

START-OF-SELECTION.
  TRY.
    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->add_new_worksheet( ).
    
    " Test basic functionality
    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Test' ).
    
    WRITE: / 'Basic test passed'.
    
  CATCH zcx_excel INTO DATA(lx_excel).
    WRITE: / 'Error:', lx_excel->get_text( ).
  ENDTRY.
```

## Performance Debugging

### Identifying Bottlenecks

Use runtime analysis to identify performance issues:

```abap
" Enable runtime measurement
GET RUN TIME FIELD DATA(lv_start_time).

" Your Excel operations here
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Data' ).

GET RUN TIME FIELD DATA(lv_end_time).
DATA(lv_duration) = lv_end_time - lv_start_time.
WRITE: / 'Operation took:', lv_duration, 'microseconds'.
```

### Memory Profiling

Monitor memory usage patterns:

```abap
" Before operation
CALL FUNCTION 'MEMORY_GET_INFO'
  IMPORTING
    allocated_bytes = DATA(lv_memory_before).

" Perform Excel operations
" ...

" After operation  
CALL FUNCTION 'MEMORY_GET_INFO'
  IMPORTING
    allocated_bytes = DATA(lv_memory_after).

DATA(lv_memory_used) = lv_memory_after - lv_memory_before.
WRITE: / 'Memory used:', lv_memory_used, 'bytes'.
```

## Error Analysis

### Exception Handling

Implement comprehensive exception handling:

```abap
TRY.
    " Excel operations
    lo_excel->save( ).
    
  CATCH zcx_excel_found INTO DATA(lx_found).
    " Handle specific Excel exceptions
    WRITE: / 'Excel error:', lx_found->get_text( ).
    
  CATCH cx_sy_conversion_error INTO DATA(lx_conversion).
    " Handle data conversion errors
    WRITE: / 'Conversion error:', lx_conversion->get_text( ).
    
  CATCH cx_root INTO DATA(lx_root).
    " Handle any other errors
    WRITE: / 'General error:', lx_root->get_text( ).
ENDTRY.
```

## Troubleshooting Checklist

1. **Verify Installation**
   - [ ] Run `ZDEMO_EXCEL_CHECKER`
   - [ ] Check all objects are active
   - [ ] Verify required SAP notes are implemented

2. **Check System Resources**
   - [ ] Available memory
   - [ ] Temporary file space
   - [ ] User authorizations

3. **Validate Input Data**
   - [ ] Data types and formats
   - [ ] Special characters
   - [ ] Large dataset handling

4. **Test Incrementally**
   - [ ] Start with simple examples
   - [ ] Add complexity gradually
   - [ ] Isolate problematic areas
