# Unit Testing Guidelines

Guidelines for testing abap2xlsx functionality and ensuring code quality.

## Testing Philosophy

abap2xlsx follows these testing principles:

- All new features should include unit tests
- Tests should cover both positive and negative scenarios
- Use meaningful test data that reflects real-world usage

## Test Structure

### Demo Programs

The library includes comprehensive demo programs that serve as both examples and tests:

- `ZDEMO_EXCEL_CHECKER` - Main test suite that validates core functionality
- Various `ZDEMO_EXCEL*` programs for specific features

### Running Tests

Before any release, ensure `ZDEMO_EXCEL_CHECKER` shows all green checkmarks.

## Writing Tests

### Basic Test Pattern

```abap
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet.

" Setup
CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).

" Test action
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Test' ).

" Verify result
DATA(lv_value) = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 ).
ASSERT lv_value = 'Test'.
```

### Error Testing

Always test error conditions and exception handling:

```abap
TRY.
    lo_worksheet->set_cell( ip_column = 'INVALID' ip_row = 1 ip_value = 'Test' ).
    ASSERT 1 = 0. " Should not reach here
  CATCH zcx_excel.
    " Expected exception
ENDTRY.
```

## Test Coverage Areas

- Cell operations (read/write)
- Formatting and styles
- Formula calculations
- File I/O operations
- Large dataset handling
- Error conditions
