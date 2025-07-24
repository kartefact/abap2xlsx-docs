# Batch Processing Multiple Files

Examples for processing large datasets and multiple Excel files efficiently.

## Large Dataset Processing

### Chunked Data Processing

```abap
" Process large internal table in chunks
CLASS zcl_excel_batch_processor DEFINITION.
  PUBLIC SECTION.
    METHODS: process_large_dataset
               IMPORTING it_data TYPE ANY TABLE
               RETURNING VALUE(rv_excel) TYPE xstring.
  PRIVATE SECTION.
    CONSTANTS: c_chunk_size TYPE i VALUE 10000.
    METHODS: process_chunk
               IMPORTING it_chunk TYPE ANY TABLE
                         io_worksheet TYPE REF TO zcl_excel_worksheet
                         iv_start_row TYPE i.
ENDCLASS.

CLASS zcl_excel_batch_processor IMPLEMENTATION.
  METHOD process_large_dataset.
    DATA: lo_excel TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zcl_excel_writer_huge_file,
          lt_chunk TYPE TABLE OF string,
          lv_current_row TYPE i VALUE 1.

    CREATE OBJECT lo_excel.
    CREATE OBJECT lo_writer.
    lo_worksheet = lo_excel->add_new_worksheet( ).
    
    " Process data in chunks
    DATA: lv_total_rows TYPE i,
          lv_processed TYPE i.
    
    lv_total_rows = lines( it_data ).
    
    LOOP AT it_data INTO DATA(ls_data).
      APPEND ls_data TO lt_chunk.
      ADD 1 TO lv_processed.
      
      " Process chunk when size reached or at end
      IF lines( lt_chunk ) = c_chunk_size OR lv_processed = lv_total_rows.
        process_chunk( 
          it_chunk = lt_chunk
          io_worksheet = lo_worksheet
          iv_start_row = lv_current_row
        ).
        
        lv_current_row = lv_current_row + lines( lt_chunk ).
        CLEAR lt_chunk.
        
        " Optional: Force garbage collection
        CALL FUNCTION 'SYSTEM_RESET_MEMORY'.
      ENDIF.
    ENDLOOP.
    
    rv_excel = lo_writer->write_file( lo_excel ).
  ENDMETHOD.

  METHOD process_chunk.
    " Process individual chunk
    LOOP AT it_chunk INTO DATA(ls_row).
      lo_worksheet->set_cell(
        ip_column = 'A'
        ip_row = iv_start_row + sy-tabix - 1
        ip_value = ls_row
      ).
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
```

### Memory-Efficient Processing

```abap
" Memory-conscious Excel generation
REPORT zexcel_memory_efficient.

DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zcl_excel_writer_huge_file.

PARAMETERS: p_rows TYPE i DEFAULT 100000.

START-OF-SELECTION.
  " Monitor memory usage
  DATA: lv_memory_start TYPE i,
        lv_memory_current TYPE i.
  
  CALL FUNCTION 'MEMORY_GET_INFO'
    IMPORTING allocated_bytes = lv_memory_start.
  
  CREATE OBJECT lo_excel.
  CREATE OBJECT lo_writer.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  
  " Generate data in batches
  DATA: lv_batch_size TYPE i VALUE 1000,
        lv_current_row TYPE i VALUE 1.
  
  DO p_rows TIMES.
    " Add row data
    lo_worksheet->set_cell(
      ip_column = 'A'
      ip_row = lv_current_row
      ip_value = |Row { lv_current_row }|
    ).
    
    ADD 1 TO lv_current_row.
    
    " Check memory every batch
    IF lv_current_row MOD lv_batch_size = 0.
      CALL FUNCTION 'MEMORY_GET_INFO'
        IMPORTING allocated_bytes = lv_memory_current.
      
      DATA(lv_memory_used) = lv_memory_current - lv_memory_start.
      WRITE: / |Processed { lv_current_row } rows, Memory: { lv_memory_used } bytes|.
      
      " Optional: Trigger garbage collection
      IF lv_memory_used > 100000000. " 100MB threshold
        CALL FUNCTION 'SYSTEM_RESET_MEMORY'.
      ENDIF.
    ENDIF.
  ENDDO.
  
  " Write final file
  DATA(lv_excel_data) = lo_writer->write_file( lo_excel ).
  WRITE: / |Final file size: { xstrlen( lv_excel_data ) } bytes|.
```

## Multiple File Generation

### Parallel File Processing

```abap
" Generate multiple Excel files in parallel
CLASS zcl_parallel_excel_generator DEFINITION.
  PUBLIC SECTION.
    TYPES: BEGIN OF ty_file_request,
             filename TYPE string,
             data TYPE REF TO data,
           END OF ty_file_request,
           tt_file_requests TYPE TABLE OF ty_file_request.
    
    METHODS: generate_multiple_files
               IMPORTING it_requests TYPE tt_file_requests
               RETURNING VALUE(rt_results) TYPE string_table.
ENDCLASS.

CLASS zcl_parallel_excel_generator IMPLEMENTATION.
  METHOD generate_multiple_files.
    " Process each file request
    LOOP AT it_requests INTO DATA(ls_request).
      " Generate individual Excel file
      DATA: lo_excel TYPE REF TO zcl_excel,
            lo_worksheet TYPE REF TO zcl_excel_worksheet,
            lo_writer TYPE REF TO zif_excel_writer.
      
      CREATE OBJECT lo_excel.
      CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
      lo_worksheet = lo_excel->add_new_worksheet( ).
      
      " Process data for this file
      FIELD-SYMBOLS: <lt_data> TYPE STANDARD TABLE.
      ASSIGN ls_request-data->* TO <lt_data>.
      
      " Add data to worksheet
      LOOP AT <lt_data> INTO DATA(ls_data).
        " Implementation specific to data structure
      ENDLOOP.
      
      " Save file
      DATA(lv_excel_data) = lo_writer->write_file( lo_excel ).
      
      " Store result
      APPEND |File { ls_request-filename } generated successfully| TO rt_results.
    ENDLOOP.
  ENDMETHOD.
ENDCLASS.
```

### Background Job Coordination

```abap
" Coordinate multiple background jobs for Excel generation
REPORT zexcel_job_coordinator.

PARAMETERS: p_jobs TYPE i DEFAULT 5.

START-OF-SELECTION.
  " Submit multiple background jobs
  DATA: lv_job_name TYPE btcjob,
        lv_job_number TYPE btcjobcnt.
  
  DO p_jobs TIMES.
    " Create unique job name
    lv_job_name = |EXCEL_GEN_{ sy-index }|.
    
    " Submit background job
    SUBMIT zexcel_background_worker
      WITH p_job_id = sy-index
      VIA JOB lv_job_name
      NUMBER lv_job_number
      AND RETURN.
    
    WRITE: / |Submitted job { lv_job_name } with number { lv_job_number }|.
  ENDDO.
  
  " Monitor job completion
  PERFORM monitor_jobs.

FORM monitor_jobs.
  " Implementation to check job status
  " and collect results when complete
ENDFORM.
```

## Performance Optimization

### Efficient Data Structures

```abap
" Use efficient data structures for batch processing
DATA: BEGIN OF ls_optimized_data,
        field1 TYPE string,
        field2 TYPE i,
        field3 TYPE p DECIMALS 2,
      END OF ls_optimized_data,
      lt_optimized_data LIKE TABLE OF ls_optimized_data.

" Avoid deep structures and object references in large tables
" Use appropriate data types to minimize memory usage
```

### Streaming Data Processing

```abap
" Stream data directly to Excel without storing in memory
CLASS zcl_excel_streamer DEFINITION.
  PUBLIC SECTION.
    METHODS: stream_data_to_excel
               IMPORTING iv_source TYPE string
               RETURNING VALUE(rv_excel) TYPE xstring.
  PRIVATE SECTION.
    METHODS: read_data_chunk
               IMPORTING iv_offset TYPE i
                         iv_size TYPE i
               RETURNING VALUE(rt_data) TYPE string_table.
ENDCLASS.

CLASS zcl_excel_streamer IMPLEMENTATION.
  METHOD stream_data_to_excel.
    DATA: lo_excel TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_writer TYPE REF TO zcl_excel_writer_huge_file.
    
    CREATE OBJECT lo_excel.
    CREATE OBJECT lo_writer.
    lo_worksheet = lo_excel->add_new_worksheet( ).
    
    " Stream data in chunks
    DATA: lv_offset TYPE i VALUE 0,
          lv_chunk_size TYPE i VALUE 1000,
          lv_current_row TYPE i VALUE 1.
    
    DO.
      " Read next chunk
      DATA(lt_chunk) = read_data_chunk( 
        iv_offset = lv_offset
        iv_size = lv_chunk_size
      ).
      
      " Exit if no more data
      IF lines( lt_chunk ) = 0.
        EXIT.
      ENDIF.
      
      " Process chunk
      LOOP AT lt_chunk INTO DATA(lv_data).
        lo_worksheet->set_cell(
          ip_column = 'A'
          ip_row = lv_current_row
          ip_value = lv_data
        ).
        ADD 1 TO lv_current_row.
      ENDLOOP.
      
      " Update offset for next chunk
      lv_offset = lv_offset + lv_chunk_size.
    ENDDO.
    
    rv_excel = lo_writer->write_file( lo_excel ).
  ENDMETHOD.

  METHOD read_data_chunk.
    " Implementation to read data chunk from source
    " Could be database, file system, or other data source
  ENDMETHOD.
ENDCLASS.
```
