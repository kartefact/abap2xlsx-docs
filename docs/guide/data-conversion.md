# Data Conversion

This guide covers the different approaches for converting ABAP data structures into Excel format and back, with a focus on performance, flexibility, and cloud compatibility.

## Overview of Conversion Methods

Abap2xlsx provides several mechanisms to populate worksheets from ABAP data:

| Method | Best For | Cloud Safe |
|---|---|---|
| `bind_table()` | Simple, flat internal tables | ✅ |
| `zcl_excel_converter` | Flexible field-level mapping | ✅ |
| Cell-by-cell (`set_cell`) | Full control, custom layouts | ✅ |
| ALV integration | Reusing ALV field catalogs | ❌ (on-premise only) |

## bind_table — Simple Table Export

The fastest way to export a flat internal table to a worksheet:

```abap
DATA: lt_data TYPE TABLE OF sflight.
SELECT * FROM sflight INTO TABLE @lt_data UP TO 100 ROWS.

lo_worksheet->bind_table(
  ip_table       = lt_data
  it_field_names = VALUE #(
    ( columnname = 'CARRID'   text = 'Carrier' )
    ( columnname = 'CONNID'   text = 'Connection' )
    ( columnname = 'FLDATE'   text = 'Flight Date' )
    ( columnname = 'PRICE'    text = 'Price' )
  )
).
```

`bind_table` generates one column per field. Use `it_field_names` to control which fields are included and to set human-readable column headers.

## zcl_excel_converter — Field-Level Mapping

`zcl_excel_converter` maps each field of a structure to a specific cell column. This gives you full control over output position and formatting without writing a cell-by-cell loop.

### Basic Usage

```abap
DATA: lo_converter TYPE REF TO zcl_excel_converter,
      lt_mapping   TYPE zcl_excel_converter=>tt_mapping.

CREATE OBJECT lo_converter.

" Build field-to-column mapping
lt_mapping = VALUE #(
  ( fieldname = 'CARRID' column = 1 )
  ( fieldname = 'CONNID' column = 2 )
  ( fieldname = 'FLDATE' column = 3 )
  ( fieldname = 'PRICE'  column = 4 )
).

lo_converter->convert(
  EXPORTING
    it_data    = lt_data
    it_mapping = lt_mapping
  CHANGING
    co_worksheet = lo_worksheet
).
```

### LOOP_NORMAL Fix — Correct Dynamic ASSIGN Handling (Jun 2025)

> **Bug fixed in Jun 2025 (PR [#1310](https://github.com/abap2xlsx/abap2xlsx/pull/1310)):** The `LOOP_NORMAL` internal method of `zcl_excel_converter` contained two defects:
>
> 1. After a dynamic `ASSIGN` (`ASSIGN COMPONENT ... OF STRUCTURE ... TO <fs>`), the code tested `IF <fs> IS ASSIGNED` — which always returns `TRUE` after a successful assignment but does **not** detect a failed one. The correct check is `IF sy-subrc = 0`.
> 2. When two invalid column names were supplied simultaneously, only the first error was raised. Both are now caught and reported correctly.

If you were calling `zcl_excel_converter` with field names derived from dynamic logic and noticing silently missing columns, update to the latest abap2xlsx version. No changes to calling code are required — the fix is internal to the converter class.

### Handling Invalid Field Names

With the fix in place, the converter correctly raises a `zcx_excel` exception when a field name in the mapping does not exist in the data structure:

```abap
TRY.
    lo_converter->convert(
      EXPORTING
        it_data    = lt_data
        it_mapping = lt_mapping_with_typo  " contains a bad field name
      CHANGING
        co_worksheet = lo_worksheet
    ).
  CATCH zcx_excel INTO DATA(lx_excel).
    " Now reliably raised for ALL invalid field names
    MESSAGE |Converter error: { lx_excel->get_text( ) }| TYPE 'E'.
ENDTRY.
```

## Cell-by-Cell Population

For maximum layout control — custom headers, merged cells, non-tabular layouts:

```abap
" Write header row manually
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Carrier' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Connection' ).
lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Date' ).
lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Price' ).

" Write data rows
DATA(lv_row) = 2.
LOOP AT lt_data INTO DATA(ls_flight).
  lo_worksheet->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_flight-carrid ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_flight-connid ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_flight-fldate ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = lv_row ip_value = ls_flight-price ).
  ADD 1 TO lv_row.
ENDLOOP.
```

## Reading Data Back

See the dedicated **[Reading Excel](/guide/reading-excel)** guide for full coverage of `zcl_excel_reader_2007`, including:

- Reading table `totalsRowFunction` attributes (Feb 2026)
- SAP Note 2922674 XML namespace fix in reader (Nov 2025)
- `get_style_from_guid` now public (Jun 2025)

## ALV Integration (On-Premise Only)

If you already have an ALV field catalog, abap2xlsx can reuse it directly. See **[ALV Integration](/guide/alv-integration)**.

> **Cloud note:** ALV-based conversion classes live in `src/not_cloud/` and are not available in S/4HANA Cloud or BTP ABAP Environment. Use `bind_table()` or cell-by-cell population instead. See **[Cloud Compatibility](/guide/cloud-compatibility)**.

## Data Type Handling

### Dates

Abap2xlsx automatically converts ABAP `d` typed fields to Excel date serial numbers when using `bind_table` or `set_cell`. Use `zcl_excel_common=>excel_string_to_date()` for manual conversions when reading back.

### Amounts and Quantities

For currency-typed fields (`curr`, `quan`), apply a number format style to ensure Excel displays the correct decimal places:

```abap
DATA: lo_style     TYPE REF TO zcl_excel_style,
      lo_num_format TYPE REF TO zcl_excel_style_number_format.

lo_style     = lo_excel->add_new_style( ).
lo_num_format = lo_style->number_format.
lo_num_format->set_format_code( '#,##0.00' ).

lo_worksheet->set_cell(
  ip_column = 'D'
  ip_row    = lv_row
  ip_value  = ls_flight-price
  ip_style  = lo_style
).
```

## Next Steps

- **[Worksheets](/guide/worksheets)** — Manage multiple sheets and comment box positioning
- **[Formatting](/guide/formatting)** — Apply styles to exported data
- **[Reading Excel](/guide/reading-excel)** — Round-trip your data
- **[Cloud Compatibility](/guide/cloud-compatibility)** — Restrictions for BTP/S/4HANA Cloud
- **[Performance](/guide/performance)** — Tips for large datasets
- **[Changelog](/guide/changelog)** — Full history of recent changes
