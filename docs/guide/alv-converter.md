# ALV / SALV Converter and OLE Automation

The classes described on this page are in the **`not_cloud`** package — they depend on
classic ABAP APIs that are not available on SAP BTP ABAP Environment or S/4HANA Public
Cloud. They install and run on:

- SAP ECC 6.0 (all Enhancement Packages)
- SAP S/4HANA On-Premise
- SAP S/4HANA Private Cloud

If you are targeting **BTP ABAP** or **S/4HANA Public Cloud**, skip to
**[Cloud Compatibility](/guide/cloud-compatibility)** and use the cloud-safe alternatives.

## Package Overview

| Class | Purpose |
|---|---|
| `zcl_excel_converter` | Core converter — maps ABAP internal-table + field-catalogue to an Excel workbook |
| `zcl_excel_converter_alv` | Converts an ALV `LVC_S_FCAT` field catalogue to Excel |
| `zcl_excel_converter_alv_grid` | Adaptor that reads field catalogue from a live `CL_GUI_ALV_GRID` instance |
| `zcl_excel_converter_salv_table` | Converts a `CL_SALV_TABLE` model to Excel, inheriting from `_alv` |
| `zcl_excel_converter_salv_model` | Helper — extracts column metadata from `CL_SALV_TABLE` |
| `zcl_excel_converter_result` | Base result class — carries the produced `zcl_excel` object |
| `zcl_excel_converter_result_ex` | Extended result — also carries exception information |
| `zcl_excel_converter_result_wd` | Web Dynpro result — delivers the file to a `CL_WD_RUNTIME_SERVICES` context |
| `zcl_excel_ole` | OLE automation — controls Excel directly on the SAP front-end GUI |
| `zexcel_template_get_types` | Report that inspects ABAP data-object types for template filling |

## ALV → Excel: `zcl_excel_converter_alv`

The quickest way to export an existing ALV/SALV report to Excel:

```abap
" Export a SALV table to Excel
DATA: lo_salv      TYPE REF TO cl_salv_table,
      lo_converter TYPE REF TO zcl_excel_converter_salv_table,
      lo_result    TYPE REF TO zcl_excel_converter_result.

" Assume lo_salv was created by the normal CL_SALV_TABLE=>FACTORY call

CREATE OBJECT lo_converter.
lo_result = lo_converter->convert(
  io_salv_table = lo_salv
  it_data       = lt_data          " The underlying ABAP internal table
  iv_sheet_name = 'Report'         " Optional tab label
).

" Retrieve the workbook and write to file
DATA(lo_excel) = lo_result->get_excel( ).
DATA(lo_writer) = NEW zcl_excel_writer_2007( ).
DATA(lv_file) = lo_writer->write_file( lo_excel ).

" Download to front-end — example using ABAP2XLSX helper
cl_gui_frontend_services=>gui_download(
  EXPORTING
    filename   = 'report.xlsx'
    filetype   = 'BIN'
  CHANGING
    data_tab   = VALUE filetable( )
  EXCEPTIONS OTHERS = 1
).
```

## Generic Converter: `zcl_excel_converter`

For precise control, use the base `zcl_excel_converter` class with a manually built
field catalogue (`zexcel_s_converter_fcat`):

```abap
DATA: lo_converter TYPE REF TO zcl_excel_converter,
      lt_fcat      TYPE zexcel_t_converter_fcat.

" Build field catalogue
APPEND VALUE zexcel_s_converter_fcat(
  fieldname  = 'MATNR'
  coltext    = 'Material'
  col_pos    = 1
  outputlen  = 18
  datatype   = 'CHAR'
) TO lt_fcat.

APPEND VALUE zexcel_s_converter_fcat(
  fieldname  = 'MAKTX'
  coltext    = 'Description'
  col_pos    = 2
  outputlen  = 40
  datatype   = 'CHAR'
) TO lt_fcat.

APPEND VALUE zexcel_s_converter_fcat(
  fieldname  = 'LABST'
  coltext    = 'Unrestricted Stock'
  col_pos    = 3
  outputlen  = 13
  datatype   = 'QUAN'
  cfieldname = 'MEINS'   " Currency/unit reference field
) TO lt_fcat.

CREATE OBJECT lo_converter.
DATA(lo_result) = lo_converter->convert(
  it_data  = lt_material_data
  it_fcat  = lt_fcat
).
```

### `zexcel_s_converter_fcat` Key Fields

| Field | Type | Description |
|---|---|---|
| `fieldname` | `FIELDNAME` | ABAP component name in the data table |
| `coltext` | `string` | Column header label |
| `col_pos` | `i` | Column order (1-based) |
| `no_out` | `flag` | `'X'` = exclude this field from output |
| `outputlen` | `intlen` | Column width (character units) |
| `datatype` | `DOMNAME` | ABAP data type (`CHAR`, `NUMC`, `CURR`, `QUAN`, `DATS`, `TIMS`, …) |
| `cfieldname` | `FIELDNAME` | Companion field for currency/unit (maps to the `ip_currency` parameter of `set_cell`) |
| `convexit` | `CONVEXIT` | Conversion exit to apply (e.g. `MATN1` for material numbers) |
| `key` | `flag` | `'X'` = mark as key column (bold by default) |
| `ref_table` / `ref_field` | `TABNAME` / `FIELDNAME` | DDIC reference for F4 help and data-type derivation |

## Web Dynpro: `zcl_excel_converter_result_wd`

For Web Dynpro ABAP applications, use the `_result_wd` subclass to push the file directly
to the browser via the WD file download UI element:

```abap
" In a WD action handler
DATA(lo_result_wd) = CAST zcl_excel_converter_result_wd(
  lo_converter->convert( it_data = lt_data  it_fcat = lt_fcat )
).

lo_result_wd->download(
  iv_filename = 'export.xlsx'
  io_component = wd_this->wd_get_api( )
).
```

## OLE Automation: `zcl_excel_ole`

`zcl_excel_ole` drives Excel directly on the SAP GUI front-end using OLE/ActiveX. This is
the legacy approach for scenarios where the server-side OOXML approach is not possible,
or where you need features like executing macros, printing via the Excel print driver, or
accessing COM add-ins.

> ⚠️ **OLE automation requires** a Windows front-end with Excel installed. It does not
> work in SAP GUI for HTML, SAP Fiori, or any browser-based UI. It also does not work in
> background jobs (no GUI session).

```abap
DATA: lo_ole TYPE REF TO zcl_excel_ole.

CREATE OBJECT lo_ole.

" Open or create a workbook on the front-end
lo_ole->open_workbook( iv_filename = 'C:\Temp\report.xlsx' ).

" Write a cell value via OLE
lo_ole->set_cell(
  iv_sheet_name = 'Sheet1'
  iv_column     = 2
  iv_row        = 5
  iv_value      = 'Hello from ABAP'
).

" Save and close
lo_ole->save_workbook( ).
lo_ole->close_workbook( ).
lo_ole->destroy( ).  " Always call destroy to release COM objects
```

### When to Use OLE vs. Server-Side API

| Criterion | Server-side (`zcl_excel_writer_*`) | OLE (`zcl_excel_ole`) |
|---|---|---|
| Background job | ✅ | ❌ |
| SAP Fiori / Web GUI | ✅ | ❌ |
| Cloud / BTP | ✅ | ❌ |
| Execute VBA macros | ❌ | ✅ |
| Print via Excel driver | ❌ | ✅ |
| COM add-in interaction | ❌ | ✅ |
| Performance (large files) | ✅ Fast | ⚠️ Slow (OLE overhead) |

## Template Helper: `zexcel_template_get_types`

`zexcel_template_get_types` is an executable ABAP report (`SE38`) that inspects an ABAP
data object and displays the component-to-range mapping for use with
`zcl_excel_fill_template`. Run it in SE38 before designing a template to understand how
your ABAP structures will map to named ranges:

1. Open SE38, enter `ZEXCEL_TEMPLATE_GET_TYPES`, and execute.
2. Enter the name of the ABAP structure or table type.
3. The report lists all components with their data types, nesting levels, and the corresponding
   named-range name that `zcl_excel_fill_template` will expect.

This is especially useful for deeply nested structures used in multi-level repeating ranges.

## Next Steps

- **[Template Filling](/guide/template-filling)** — server-side template filling with named ranges
- **[Cloud Compatibility](/guide/cloud-compatibility)** — cloud-safe alternatives to `not_cloud` classes
- **[Data Conversion](/guide/data-conversion)** — `bind_table` as a simpler alternative
- **[XLSM Macros](/guide/xlsm-macros)** — preserve VBA macros in a server-side workflow
