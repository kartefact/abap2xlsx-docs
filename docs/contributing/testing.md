# Testing Guide

Guidelines for testing abap2xlsx functionality and ensuring code quality.

## Testing Philosophy

abap2xlsx follows these testing principles:

- All new features should include ABAP Unit tests where feasible
- Tests should cover both positive and negative scenarios
- Use meaningful test data that reflects real-world usage
- Regression tests are especially important for reader/writer round-trip changes

## Test Structure

### Demo Programs (`ZDEMO_EXCEL*`)

The library ships a comprehensive set of demo programs that serve as both
working examples and an integration test suite:

| Program | Purpose |
|---|---|
| `ZDEMO_EXCEL_CHECKER` | Runs all demos in batch and verifies no short dump or exception occurs |
| `ZDEMO_EXCEL01` … `ZDEMO_EXCELnn` | Individual feature demos (formatting, formulas, images, charts, …) |
| `ZDEMO_EXCEL_ALV*` | ALV-to-Excel conversion demos (on-premise only) |

### ABAP Unit Tests

Selected classes contain local test classes using ABAP Unit:

- `ZCL_EXCEL` — style registry, `get_style_from_guid` (added Jun 2025, PR
  [#1315](https://github.com/abap2xlsx/abap2xlsx/pull/1315))
- `ZCL_EXCEL_WORKSHEET` — `check_rtf` logic (added Jun 2025, PR [#1315](https://github.com/abap2xlsx/abap2xlsx/pull/1315))
- `ZCL_EXCEL_COMMON` — column-to-alpha conversion helpers

Run all unit tests in SE80 / Eclipse ADT using *Run → ABAP Unit Tests* on the
`ZABAP2XLSX` package.

## Running `ZDEMO_EXCEL_CHECKER`

`ZDEMO_EXCEL_CHECKER` is the primary regression gate. It executes every demo
program sequentially and reports pass/fail per demo.

### Before any Release

Ensure `ZDEMO_EXCEL_CHECKER` shows **all green checkmarks** before tagging a
release. A failing demo indicates either a regression in the library or a demo
that must be updated.

### Running on Older / Restricted Systems

> **From the automated-tests.md added in PR [#1342](https://github.com/abap2xlsx/abap2xlsx/pull/1342):**

On older NetWeaver releases or restricted sandbox systems you may encounter the
following situations:

- **Missing `ZDEMO_EXCEL_CHECKER`** — the checker program was added later than
  some of the demos. Install it manually from the repository before running.
- **Authorization issues** — the checker calls `SUBMIT ... AND RETURN`. Ensure
  the test user has `S_PROGRAM` authorization for all `ZDEMO_*` programs.
- **Syntax errors on activation** — a handful of demos use inline declarations
  (`DATA(lv_x) = ...`) that require kernel ≥ 7.40. Exclude affected demos from
  the checker's selection or apply the syntax-compat fix from PR
  [#1344](https://github.com/abap2xlsx/abap2xlsx/pull/1344).
- **ALV demos fail on cloud** — `ZDEMO_EXCEL_ALV*` programs reside in
  `src/not_cloud/`. Skip them when running on BTP ABAP Environment or
  S/4HANA Cloud. Filter by excluding programs matching `ZDEMO_EXCEL_ALV*` in
  the checker's selection screen.
- **File-system write errors** — some demos write output files to the
  application server. Confirm the work directory is writable, or use the
  SAP GUI download option instead.

### Selection Screen Parameters

The checker's selection screen lets you narrow which demos are executed:

```
Program range:  ZDEMO_EXCEL* (default — runs everything)
Exclude range:  ZDEMO_EXCEL_ALV* (recommended on cloud / restricted systems)
```

## Writing New Tests

### Basic ABAP Unit Test Pattern

```abap
CLASS ltcl_my_feature DEFINITION FOR TESTING
  RISK LEVEL HARMLESS DURATION SHORT.

  PRIVATE SECTION.
    METHODS: test_set_and_get FOR TESTING.
ENDCLASS.

CLASS ltcl_my_feature IMPLEMENTATION.
  METHOD test_set_and_get.
    DATA: lo_excel     TYPE REF TO zcl_excel,
          lo_worksheet TYPE REF TO zcl_excel_worksheet.

    CREATE OBJECT lo_excel.
    lo_worksheet = lo_excel->add_new_worksheet( ).

    lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Hello' ).

    cl_abap_unit_assert=>assert_equals(
      act = lo_worksheet->get_cell( ip_column = 'A' ip_row = 1 )
      exp = 'Hello'
      msg = 'set_cell / get_cell round-trip'
    ).
  ENDMETHOD.
ENDCLASS.
```

### Error / Exception Testing

Always verify that invalid inputs raise `zcx_excel`:

```abap
METHOD test_invalid_column.
  TRY.
      lo_worksheet->set_cell( ip_column = '' ip_row = 1 ip_value = 'X' ).
      cl_abap_unit_assert=>fail( 'Expected zcx_excel was not raised' ).
    CATCH zcx_excel.
      " Expected — pass
  ENDTRY.
ENDMETHOD.
```

### Reader Round-Trip Pattern

For reader changes (e.g. the `totalsRowFunction` fix in PR [#1296](https://github.com/abap2xlsx/abap2xlsx/pull/1296)
or the SAP Note 2922674 namespace fix in PR [#1349](https://github.com/abap2xlsx/abap2xlsx/pull/1349)),
use a write-then-read round-trip test:

```abap
METHOD test_table_totals_roundtrip.
  " 1. Build and write a workbook with a table + totals row
  DATA(lo_excel_out) = NEW zcl_excel( ).
  DATA(lo_ws_out)    = lo_excel_out->add_new_worksheet( ).
  " ... set up table with totalsRowFunction = 'sum' ...
  DATA(lo_writer) = NEW zcl_excel_writer_2007( ).
  DATA(lv_xstr)   = lo_writer->write_file( lo_excel_out ).

  " 2. Read the file back
  DATA(lo_reader)   = NEW zcl_excel_reader_2007( ).
  DATA(lo_excel_in) = lo_reader->load_file( lv_xstr ).

  " 3. Assert the totalsRowFunction survived the round-trip
  DATA(lo_ws_in)  = lo_excel_in->get_active_worksheet( ).
  DATA(lo_tables) = lo_ws_in->get_tables( ).
  DATA(lo_table)  = lo_tables->get( 1 ).
  DATA(lo_cols)   = lo_table->get_table_columns( ).
  DATA(lo_col)    = lo_cols->get( 1 ).

  cl_abap_unit_assert=>assert_equals(
    act = lo_col->get_totals_row_function( )
    exp = 'sum'
    msg = 'totalsRowFunction must survive writer → reader round-trip'
  ).
ENDMETHOD.
```

## Test Coverage Areas

- Cell operations (read/write, all supported ABAP types including `XSTRING`)
- Formatting and styles (`get_style_from_guid`, `check_rtf`)
- Formula calculations and formula preservation
- Table round-trips (totals row function, column definitions)
- Comment box geometry (`ms_box` structure, `mc_box_default` constant)
- Worksheet copy constructor (comment collection propagation)
- Reader XML namespace handling (SAP Note 2922674 scenarios)
- File I/O operations
- Large dataset handling
- Error conditions and exception classes
- Cloud-compatible syntax (no use of `DESCRIBE TABLE LINES`, `LANG`, `SEOCLSNAME`, etc.)

## Related Resources

- [automated-tests.md in the abap2xlsx source repo](https://github.com/abap2xlsx/abap2xlsx/blob/main/automated-tests.md)
  (added Oct 2025, PR [#1342](https://github.com/abap2xlsx/abap2xlsx/pull/1342))
- **[Contributing: Coding Guidelines](/contributing/coding-guidelines)**
- **[Contributing: Development Setup](/contributing/development-setup)**
- **[Changelog](/guide/changelog)** — history of test-related changes
