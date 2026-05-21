# Data Conversion

This guide covers the different approaches for converting ABAP data structures into Excel format and back.

## Overview of Conversion Methods

| Method | Best For | Cloud Safe |
|---|---|---|
| `bind_table()` | Simple, flat internal tables | Yes |
| `zcl_excel_converter` | Flexible field-level mapping | Yes |
| Cell-by-cell (`set_cell`) | Full control, custom layouts | Yes |
| ALV integration | Reusing ALV field catalogs | No (on-premise only) |

## bind_table - Simple Table Export

```abap
DATA: lt_data TYPE TABLE OF sflight.
SELECT * FROM sflight INTO TABLE @lt_data UP TO 100 ROWS.

lo_worksheet->bind_table(
  ip_table       = lt_data
  it_field_names = VALUE #(
    ( columnname = 'CARRID' text = 'Carrier' )
    ( columnname = 'CONNID' text = 'Connection' )
    ( columnname = 'FLDATE' text = 'Flight Date' )
    ( columnname = 'PRICE'  text = 'Price' )
  )
).
```

## zcl_excel_converter - Field-Level Mapping

### Basic Usage

```abap
DATA: lo_converter TYPE REF TO zcl_excel_converter,
      lt_mapping   TYPE zcl_excel_converter=>tt_mapping.

CREATE OBJECT lo_converter.

lt_mapping = VALUE #(
  ( fieldname = 'CARRID' column = 1 )
  ( fieldname = 'CONNID' column = 2 )
  ( fieldname = 'FLDATE' column = 3 )
  ( fieldname = 'PRICE'  column = 4 )
).

lo_converter->convert(
  EXPORTING it_data = lt_data  it_mapping = lt_mapping
  CHANGING  co_worksheet = lo_worksheet
).
```

### Performance Note - Pass-by-Reference (Feb 2025)

> **Changed in Feb-Mar 2025 (PR [#1037](https://github.com/abap2xlsx/abap2xlsx/pull/1037), PR [#1039](https://github.com/abap2xlsx/abap2xlsx/pull/1039)):** Large structure and table parameters previously passed by value (`VALUE`) are now passed by reference (`REFERENCE`).

This is a **transparent internal change** - your calling code does not need modification. It significantly reduces memory copying overhead for large datasets. If you have custom subclasses or method redefinitions of converter methods, review your signatures for compatibility.

### Autofilter Fix (Sep 2024)

> **Fixed in Sep 2024 (PR [#1239](https://github.com/abap2xlsx/abap2xlsx/pull/1239)):** Using `zcl_excel_converter` with an ALV field catalog that included autofilter could produce an incorrect or missing autofilter range in the output file.

Updating to the latest abap2xlsx version resolves this. No code changes required.

### LOOP_NORMAL Fix - Correct Dynamic ASSIGN Handling (Jun 2025)

> **Bug fixed in Jun 2025 (PR [#1310](https://github.com/abap2xlsx/abap2xlsx/pull/1310)):** Two defects in the `LOOP_NORMAL` internal method:
>
> 1. After a dynamic `ASSIGN`, the code tested `IF <fs> IS ASSIGNED` - which is always `TRUE` after any assignment but does **not** detect a failed one. Corrected to `IF sy-subrc = 0`.
> 2. When two invalid column names were supplied simultaneously, only the first error was raised. Both are now caught.

```abap
TRY.
    lo_converter->convert(
      EXPORTING it_data = lt_data  it_mapping = lt_mapping_with_typo
      CHANGING  co_worksheet = lo_worksheet
    ).
  CATCH zcx_excel INTO DATA(lx_excel).
    " Now reliably raised for ALL invalid field names
    MESSAGE |Converter error: { lx_excel->get_text( ) }| TYPE 'E'.
ENDTRY.
```

## Cell-by-Cell Population

```abap
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Carrier' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Connection' ).

DATA(lv_row) = 2.
LOOP AT lt_data INTO DATA(ls_flight).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_flight-carrid ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_flight-connid ).
  ADD 1 TO lv_row.
ENDLOOP.
```

## Reading Data Back

See **[Reading Excel](/guide/reading-excel)** for full coverage of `zcl_excel_reader_2007`, including `totalsRowFunction` round-trip (Feb 2026) and the SAP Note 2922674 XML namespace fix in the reader (Nov 2025).

## ALV Integration (On-Premise Only)

See **[ALV Integration](/guide/alv-integration)**.

> **Cloud note:** ALV classes live in `src/not_cloud/` and are unavailable in S/4HANA Cloud / BTP. Use `bind_table()` or cell-by-cell population instead. See **[Cloud Compatibility](/guide/cloud-compatibility)**.

## Data Type Handling

### Dates

Abap2xlsx automatically converts ABAP `d` typed fields to Excel date serial numbers in `bind_table` and `set_cell`. Use `zcl_excel_common=>excel_string_to_date()` for manual conversions when reading back.

### Amounts and Quantities

```abap
DATA: lo_style     TYPE REF TO zcl_excel_style,
      lo_num_format TYPE REF TO zcl_excel_style_number_format.

lo_style      = lo_excel->add_new_style( ).
lo_num_format = lo_style->number_format.
lo_num_format->set_format_code( '#,##0.00' ).

lo_worksheet->set_cell(
  ip_column = 'D'  ip_row = lv_row
  ip_value  = ls_flight-price  ip_style = lo_style
).
```

## Next Steps

- **[Worksheets](/guide/worksheets)** - Multiple sheets and comment box positioning
- **[Formatting](/guide/formatting)** - Apply styles to exported data
- **[Reading Excel](/guide/reading-excel)** - Round-trip your data
- **[CSV Export](/guide/csv-export)** - Export to CSV with skip-hidden-rows/columns
- **[Cloud Compatibility](/guide/cloud-compatibility)** - Restrictions for BTP/S/4HANA Cloud
- **[Performance](/guide/performance)** - Tips for large datasets
- **[Changelog](/guide/changelog)** - Full history of recent changes
