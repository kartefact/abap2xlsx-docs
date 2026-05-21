# OLE / On-Screen Converter (`zcl_excel_ole`)

`zcl_excel_ole` is a **non-cloud** class (located in `src/not_cloud/`) that drives a local Microsoft Excel installation via OLE Automation (ActiveX). It converts an xlsx binary produced by abap2xlsx into other formats that Excel itself supports — most commonly PDF — by opening the file in Excel and calling `SaveAs` or `ExportAsFixedFormat` programmatically.

::: warning On-screen / frontend only
`zcl_excel_ole` runs entirely on the SAP GUI frontend machine. It requires a licensed copy of Microsoft Excel to be installed on that machine and **cannot be used** on application servers, background jobs, or SAP BTP. For cloud-safe PDF generation consider the [ALV Converter](./alv-converter.md) or a third-party PDF library.
:::

## When to Use It

| Scenario | Recommended approach |
|---|---|
| GUI frontend, Excel installed, need PDF | `zcl_excel_ole` |
| Background job / batch | Not possible — use a server-side PDF tool |
| SAP BTP / steampunk | Not possible — `src/not_cloud` excluded |
| Need xlsx only | Use standard `zcl_excel_writer_2007` |

## Package Location

`zcl_excel_ole` lives in the sub-package `src/not_cloud/`. Install this package only if your system is a classic on-premise ABAP stack and you need OLE features. Cloud-ready deployments should exclude the `not_cloud` folder entirely (it is deliberately separated for this reason).

## Basic Usage Pattern

```abap
" 1. Build an excel object and produce the xstring as usual
DATA lo_excel  TYPE REF TO zcl_excel.
DATA lo_writer TYPE REF TO zcl_excel_writer_2007.
DATA lv_xstring TYPE xstring.

lo_excel  = NEW zcl_excel( ).
" ... populate worksheets ...
lo_writer = NEW zcl_excel_writer_2007( ).
lv_xstring = lo_writer->write_file( lo_excel ).

" 2. Save the xstring to a local temp file (frontend)
DATA lv_path TYPE string VALUE 'C:\Temp\myreport.xlsx'.
" ... use GUI_DOWNLOAD or cl_gui_frontend_services to write the binary ...

" 3. Open in Excel via OLE and convert
DATA lo_ole TYPE REF TO zcl_excel_ole.
lo_ole = NEW zcl_excel_ole( ).
lo_ole->convert(
  iv_source_path = lv_path
  iv_target_path = 'C:\Temp\myreport.pdf'
  iv_format      = zcl_excel_ole=>c_format_pdf
).
```

::: tip Error handling
PR bug fix (2025) corrected an unhandled exception in `LOOP_NORMAL` inside `zcl_excel_converter`. Ensure you are on a commit after `3bdf779` to benefit from this fix, and always wrap OLE calls in a `TRY ... CATCH cx_sy_native_type_conflict cx_root.` block because Excel OLE automation can raise a wide variety of exceptions depending on the Excel version installed.
:::

## Converter Class Hierarchy (`not_cloud`)

The `not_cloud` package contains a small converter framework:

| Class | Purpose |
|---|---|
| `zif_excel_converter` | Interface — `convert()` method signature |
| `zcl_excel_converter` | Base implementation — drives OLE, handles normal/special loop |
| `zcl_excel_converter_alv` | Subclass — converts from an ALV model |
| `zcl_excel_converter_alv_grid` | Subclass — converts from a live `cl_gui_alv_grid` |
| `zcl_excel_converter_salv_table` | Subclass — converts from `cl_salv_table` |
| `zcl_excel_converter_result` | Result descriptor (path, format, success flag) |
| `zcl_excel_converter_result_ex` | Extended result with error details |
| `zcl_excel_converter_result_wd` | Result adapter for Web Dynpro scenarios |
| `zcl_excel_ole` | Low-level OLE driver — opens Excel, calls SaveAs |

## Format Constants

Check `zcl_excel_ole` for the latest set; typical constants include:

```abap
zcl_excel_ole=>c_format_pdf    " PDF via ExportAsFixedFormat
zcl_excel_ole=>c_format_xlsx   " re-save as xlsx (normalises file)
zcl_excel_ole=>c_format_csv    " active sheet as CSV
```

## ALV-to-Excel via OLE

`zcl_excel_converter_alv` wraps the classic ALV-to-Excel path where the data is written to an ALV grid first and then the live OLE-controlled Excel instance is used to capture it:

```abap
DATA lo_alv_converter TYPE REF TO zcl_excel_converter_alv.
lo_alv_converter = NEW zcl_excel_converter_alv( ).
lo_alv_converter->convert(
  io_alv_model    = lo_salv_model
  iv_target_path  = 'C:\Temp\alv_output.xlsx'
).
```

This is the legacy approach predating `zcl_excel_converter_salv_table`. For new development, prefer the non-OLE path documented in [ALV Integration](./alv-integration.md).

## Helper Program: `zexcel_template_get_types`

The `not_cloud` package also ships `zexcel_template_get_types` — a standalone ABAP report that introspects a running Excel instance via OLE to discover the available `XlFileFormat` constant values on the current Excel installation. Run it once on the target frontend machine to verify which format codes are available before hard-coding constants in production code.
