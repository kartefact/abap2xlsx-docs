# CSV Export

Abap2xlsx includes a dedicated CSV writer class, `zcl_excel_writer_csv`, that converts an in-memory workbook into comma-separated (or custom-delimited) text.

## Basic Usage

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_writer    TYPE REF TO zcl_excel_writer_csv.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->get_active_worksheet( ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Name' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Score' ).
lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Alice' ).
lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 95 ).

CREATE OBJECT lo_writer.
DATA(lv_csv) = lo_writer->write_file( lo_excel ).
```

The result is an `XSTRING` containing the UTF-8-encoded CSV.

## Configuring the Delimiter

```abap
" Default is comma. Change to semicolon for European locales:
lo_writer->set_separator( ';' ).
```

## Skipping Hidden Rows and Columns

> **Added in Jan 2025 (PR [#1268](https://github.com/abap2xlsx/abap2xlsx/pull/1268)).**

Two new options control whether hidden rows/columns are included in the output. This is particularly useful when exporting an ALV-based workbook where the user has hidden columns via the layout:

```abap
CREATE OBJECT lo_writer.
lo_writer->set_skip_hidden_rows( abap_true ).
lo_writer->set_skip_hidden_columns( abap_true ).
DATA(lv_csv) = lo_writer->write_file( lo_excel ).
```

### Typical ALV-to-CSV Pattern

```abap
DATA lo_csv TYPE REF TO zcl_excel_writer_csv.
CREATE OBJECT lo_csv.
lo_csv->set_skip_hidden_columns( abap_true ).
lo_csv->set_skip_hidden_rows( abap_true ).
DATA(lv_file) = lo_csv->write_file( lo_excel ).
```

## Date Format Handling

Prior to the Jan 2025 fix, a mismatch between the user logon language and the English-only domain lookup in `get_default_excel_date_format()` caused date fields to go unrecognised. PR [#1268](https://github.com/abap2xlsx/abap2xlsx/pull/1268) resolves this and also handles the NW 7.52+ domain text change:

| NW Release | Domain `XUDATFM` text example |
|---|---|
| NW 7.40 and earlier | `DD.MM.YYYY` |
| NW 7.52 and later | `DD.MM.YYYY (Gregorian Date)` |

## Notes on the Active Worksheet

The CSV writer operates on the **active worksheet only**. Set the active sheet before writing if your workbook contains multiple sheets:

```abap
lo_excel->set_active_sheet_index( 2 ).
DATA(lv_csv) = lo_csv->write_file( lo_excel ).
```

## Related Resources

- **[Basic Usage](/guide/basic-usage)** - Overview of all writer types
- **[Data Conversion](/guide/data-conversion)** - Building the workbook from ABAP data
- **[ALV Integration](/guide/alv-integration)** - Converting ALV field catalogs (on-premise)
- **[Performance](/guide/performance)** - Tips for large exports
- **[Changelog](/guide/changelog)** - Full history
