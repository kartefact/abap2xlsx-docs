# Cloud Compatibility

abap2xlsx is designed to work on **SAP BTP ABAP Environment**, **S/4HANA Public Cloud**, and
all classic on-premise ABAP stacks. This guide explains which parts of the library require
classic ABAP and which are cloud-safe.

## The `not_cloud` Package

Classes that depend on classic ABAP APIs are isolated in the **`src/not_cloud/`** package.
Do **not** install this package on cloud systems.

### `not_cloud` Class Inventory

| Class | Depends on | Alternative |
|---|---|---|
| `zcl_excel_converter` | `DDIF_FIELDINFO_GET`, `DESCRIBE TABLE LINES` | `bind_table` + `zcl_excel_worksheet` |
| `zcl_excel_converter_alv` | `LVC_*` ALV APIs | `bind_table` with manual field catalogue |
| `zcl_excel_converter_alv_grid` | `CL_GUI_ALV_GRID` | `CL_SALV_TABLE` + `zcl_excel_converter_salv_table` |
| `zcl_excel_converter_salv_table` | `CL_SALV_TABLE` model APIs | Cloud ALV (`cl_grid_display_salv`, future) |
| `zcl_excel_converter_salv_model` | `CL_SALV_*` | Same as above |
| `zcl_excel_converter_result_wd` | `CL_WD_RUNTIME_SERVICES` | OData/REST file download |
| `zcl_excel_ole` | `CL_GUI_FRONTEND_SERVICES`, OLE/COM | Server-side writer + HTTP download |
| `zexcel_template_get_types` | `DESCRIBE TABLE COMPONENTS` | Manual structure inspection |

See **[ALV Converter and OLE](/guide/alv-converter)** for full usage documentation on all
`not_cloud` classes.

## What Changed for Cloud Compatibility

The following statements were removed from the main `src/` package classes and replaced
with cloud-compatible equivalents:

| Removed statement | Replacement |
|---|---|
| `DESCRIBE TABLE lt_data LINES lv_count` | `lv_count = lines( lt_data )` |
| `DATA: lt_strtab TYPE STRTABLE` | `DATA: lt_strtab TYPE string_table` |
| `DATA: lt_sstrtab TYPE SSTRTABLE` | `DATA: lt_sstrtab TYPE string_table` |
| `DATA: lv_text TYPE SCRTEXT_M` | `DATA: lv_text TYPE string` |
| `DATA: lv_lang TYPE LANG` | `DATA: lv_lang TYPE sy-langu` |
| `DATA: lv_name TYPE SEOCLSNAME` | `DATA: lv_name TYPE string` |

The class `zcl_excel_obsolete_func_wrap` wraps function modules that are not available on
cloud (`GUI_DOWNLOAD`, `GUI_UPLOAD`, `POPUP_TO_CONFIRM`, `SAPGUI_PROGRESS_INDICATOR`). The
cloud-compatible replacements are `cl_gui_frontend_services` (on classic ABAP) or HTTP
download (on BTP).

## Cloud Installation

### abapGit

Install only the main `src/` package via abapGit. In the `.abapgit.xml`:

```xml
<asx:values>
  <DATA>
    <PACKAGE>ZABAP2XLSX</PACKAGE>
    <!-- Do NOT add ZABAP2XLSX_NOT_CLOUD for cloud systems -->
  </DATA>
</asx:values>
```

Or when cloning via abapGit online, choose only the `ZABAP2XLSX` top-level package and
**deselect** `ZABAP2XLSX_NOT_CLOUD` (or the equivalent sub-package name in your system).

### SAPlink / Manual Transport

Import only objects from `src/` — do not import anything under `src/not_cloud/`.

## Runtime Checks

A useful pattern before calling any `not_cloud` class is to check whether the class exists:

```abap
IF cl_abap_classdescr=>describe_by_name( 'ZCL_EXCEL_CONVERTER' ) IS INITIAL.
  " Running on cloud — use bind_table instead
  lo_worksheet->bind_table( ip_table = lt_data ).
ELSE.
  " Running on classic — use converter
  DATA(lo_conv) = NEW zcl_excel_converter( ).
  DATA(lo_result) = lo_conv->convert( it_data = lt_data  it_fcat = lt_fcat ).
ENDIF.
```

## Feature Matrix

| Feature | Cloud-compatible | Requires `not_cloud` |
|---|:---:|:---:|
| `zcl_excel_writer_2007` | ✅ | |
| `zcl_excel_writer_huge_file` | ✅ | |
| `zcl_excel_writer_csv` | ✅ | |
| `zcl_excel_writer_xlsm` | ✅ | |
| `zcl_excel_reader_2007` | ✅ | |
| `zcl_excel_reader_huge_file` | ✅ | |
| `zcl_excel_reader_xlsm` | ✅ | |
| `zcl_excel_fill_template` | ✅ | |
| `zcl_excel_security` (AES-256) | ✅ | |
| `zcl_excel_autofilter` | ✅ | |
| `zcl_excel_table` | ✅ | |
| `zcl_excel_style_changer` | ✅ | |
| `zcl_excel_converter` (ALV) | | ✅ |
| `zcl_excel_converter_alv` | | ✅ |
| `zcl_excel_converter_salv_table` | | ✅ |
| `zcl_excel_converter_result_wd` | | ✅ |
| `zcl_excel_ole` | | ✅ |

## Next Steps

- **[ALV Converter and OLE](/guide/alv-converter)** — full documentation for `not_cloud` classes
- **[Data Conversion](/guide/data-conversion)** — cloud-safe `bind_table` patterns
- **[Huge-File Writer](/guide/huge-file-writer)** — streaming writer for large datasets
- **[Workbook Security](/guide/workbook-security)** — AES-256 encryption
