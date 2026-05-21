# Reading Worksheet Data into an ABAP Table

`convert_to_table` is the upload counterpart to `bind_table`. Where `bind_table` writes an ABAP internal table into a worksheet, `convert_to_table` reads cell data back out of a worksheet into an ABAP internal table. It is defined on `zcl_excel_worksheet`.

## Signature

```abap
METHODS convert_to_table
  IMPORTING
    it_field_catalog TYPE zexcel_t_fieldcatalog OPTIONAL
    iv_begin_row     TYPE int4 DEFAULT 2
    iv_end_row       TYPE int4 DEFAULT 0
  EXPORTING
    et_data          TYPE STANDARD TABLE
    er_data          TYPE REF TO data
  RAISING
    zcx_excel.
```

### Parameters

| Parameter | Direction | Default | Purpose |
|---|---|---|---|
| `it_field_catalog` | Importing | — | Field catalog describing target columns and types |
| `iv_begin_row` | Importing | `2` | First data row to read (default skips a header row) |
| `iv_end_row` | Importing | `0` (all) | Last row to read; `0` means read to the last populated row |
| `et_data` | Exporting | — | Typed ABAP internal table — values are converted to the declared ABAP types |
| `er_data` | Exporting | — | `REF TO data` pointing to a string-column table — raw cell text, no conversion losses |

## Choosing between `et_data` and `er_data`

| Scenario | Use |
|---|---|
| You have a matching ABAP Dictionary structure and want typed output | `et_data` with `it_field_catalog` |
| Cell values contain leading zeros, special formatting, or precision you cannot afford to lose | `er_data` |
| You need to map only selected columns | `it_field_catalog` with `et_data` |
| Quick prototyping / unknown sheet structure | `er_data` — no catalog needed |

`et_data` and `er_data` can be used simultaneously in the same call.

## Basic example — typed read with field catalog

```abap
DATA: lo_excel     TYPE REF TO zcl_excel,
      lo_reader    TYPE REF TO zif_excel_reader,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lt_catalog   TYPE zexcel_t_fieldcatalog,
      ls_catalog   TYPE zexcel_s_fieldcatalog,
      lt_materials TYPE TABLE OF zmy_material_s.

" Read workbook
CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
lo_excel = lo_reader->load_file( lv_xstring ).
lo_worksheet = lo_excel->get_active_worksheet( ).

" Define field catalog matching column order in the sheet
ls_catalog-col_pos   = 1.  ls_catalog-fieldname = 'MATNR'.  APPEND ls_catalog TO lt_catalog.
ls_catalog-col_pos   = 2.  ls_catalog-fieldname = 'MAKTX'.  APPEND ls_catalog TO lt_catalog.
ls_catalog-col_pos   = 3.  ls_catalog-fieldname = 'MEINS'.  APPEND ls_catalog TO lt_catalog.

" Convert — row 1 is the header, data starts at row 2
lo_worksheet->convert_to_table(
  EXPORTING
    it_field_catalog = lt_catalog
    iv_begin_row     = 2
  IMPORTING
    et_data          = lt_materials ).
```

## Basic example — lossless string read

Use `er_data` when you need the raw cell text exactly as it appears in the sheet:

```abap
DATA: lr_data    TYPE REF TO data,
      lt_strings TYPE TABLE OF string.  " illustrative — actual type is dynamic

lo_worksheet->convert_to_table(
  EXPORTING
    iv_begin_row = 2
  IMPORTING
    er_data      = lr_data ).

" Dereference to access rows
ASSIGN lr_data->* TO FIELD-SYMBOL(<lt_raw>).
```

The table behind `er_data` has one string column per spreadsheet column, preserving leading zeros, date formats, and long decimals without rounding.

## Limiting the row range

```abap
" Read only rows 5 through 25
lo_worksheet->convert_to_table(
  EXPORTING
    iv_begin_row = 5
    iv_end_row   = 25
  IMPORTING
    et_data      = lt_result ).
```

Setting `iv_end_row = 0` (the default) reads all rows from `iv_begin_row` to `get_highest_row( )`.

## Field catalog for `convert_to_table`

The same `zexcel_t_fieldcatalog` / `zexcel_s_fieldcatalog` type used by `bind_table` controls the conversion:

- `col_pos` — 1-based column position in the sheet (required)
- `fieldname` — name of the target ABAP structure field
- `inttype` — ABAP type kind (optional; inferred from the target structure if omitted)
- `ref_table` / `ref_field` — dictionary reference for currency/quantity fields

When no catalog is supplied, all columns are read as strings into `er_data`.

## Relationship to `bind_table` and `get_table`

| Method | Direction | Catalog | Use case |
|---|---|---|---|
| `bind_table` | ABAP → sheet | Supported | Write typed ABAP data to a worksheet |
| `convert_to_table` | Sheet → ABAP | Supported | Read sheet data back into typed ABAP table |
| `get_table` | Sheet → ABAP | Not supported | Quick untyped extraction by row/column range |

`get_table` (parameters: `iv_skipped_rows`, `iv_skipped_cols`, `iv_max_col`, `iv_max_row`, `iv_skip_bottom_empty_rows`) is a simpler alternative for quick reads where type conversion is not needed.

## Round-trip pattern

A common integration pattern uses `bind_table` to produce a template and `convert_to_table` to ingest user-edited returns:

```abap
" 1. Export: ABAP → Excel (distribute to user)
lo_ws->bind_table( ip_table = lt_orders it_field_catalog = lt_catalog ).

" 2. User edits the Excel file and returns it
" 3. Import: Excel → ABAP
lo_ws_returned->convert_to_table(
  EXPORTING it_field_catalog = lt_catalog
  IMPORTING et_data          = lt_orders_updated ).
```

## See also

- [bind_table and Field Catalog](./excel-tables.md)
- [Reading Excel Files](./reading-excel.md)
- [Template Filling](./template-filling.md)
