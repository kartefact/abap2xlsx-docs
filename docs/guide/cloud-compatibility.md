# Cloud Compatibility (S/4HANA Cloud & BTP ABAP)

This page documents the ongoing effort to make abap2xlsx fully compatible with **SAP S/4HANA Cloud** (ABAP Environment / Steampunk) and **SAP BTP ABAP Environment**, where classic on-premise ABAP syntax is restricted.

## What Is the Cloud Compatibility Track?

Abap2xlsx has historically relied on a small number of classic ABAP language features and built-in variables that are not available in the ABAP Cloud programming model. The maintainers have been systematically removing these dependencies.

Cloud-restricted constructs that have been removed or replaced:

| Removed construct | Replaced with | PR | Merged |
|---|---|---|---|
| `DESCRIBE TABLE ... LINES` | `lines( )` built-in function | [#1340](https://github.com/abap2xlsx/abap2xlsx/pull/1340) | Jan 2026 |
| `STRTABLE` / `SSTRTABLE` type references | Standard string/table types | [#1336](https://github.com/abap2xlsx/abap2xlsx/pull/1336) | Sep 2025 |
| `SCRTEXT_S/M/L` (screen-text types) | `string` | [#1337](https://github.com/abap2xlsx/abap2xlsx/pull/1337) | Sep 2025 |
| `SEOCLSNAME` type references | `string` | [#1331](https://github.com/abap2xlsx/abap2xlsx/pull/1331) | Aug 2025 |
| `SY-LANGU` / `LANG` system field | `cl_abap_context_info=>get_system_language( )` | [#1313](https://github.com/abap2xlsx/abap2xlsx/pull/1313) | Jun 2025 |
| `IHTTPNVP` type references | `string` | [#1312](https://github.com/abap2xlsx/abap2xlsx/pull/1312) | Jun 2025 |
| `TDCWIDTHS` type references | `i` (integer) | [#1311](https://github.com/abap2xlsx/abap2xlsx/pull/1311) | May 2025 |

## Deployment Notes

### Classic On-Premise (ECC, S/4HANA On-Premise)

No changes required. All replacements above are backward-compatible with relevant on-premise ABAP releases.

### S/4HANA Cloud / BTP ABAP Environment

Abap2xlsx provides a dedicated `src/not_cloud/` sub-package. Exclude it when deploying to a cloud tenant:

```
abap2xlsx/
  src/
    zcl_excel.clas.abap            <- cloud-safe
    zcl_excel_common.clas.abap     <- cloud-safe
    ...
    not_cloud/                     <- EXCLUDE from cloud deployments
      zcl_excel_*_alv*.clas.abap
      ...
```

## Checking Your Custom Code

```abap
" Not allowed in ABAP Cloud
DESCRIBE TABLE lt_data LINES lv_count.    " use lines( lt_data )
DATA lv_lang  TYPE LANG.                  " use string or CL_ABAP_CONTEXT_INFO
DATA lv_name  TYPE SEOCLSNAME.            " use string
DATA lv_text  TYPE SCRTEXT_S.             " use string
DATA lv_http  TYPE IHTTPNVP.              " use string
DATA lv_width TYPE TDCWIDTHS.             " use i

" Cloud-compatible equivalents
DATA(lv_count) = lines( lt_data ).
DATA(lv_lang)  = cl_abap_context_info=>get_system_language( ).
DATA lv_name   TYPE string.
DATA lv_text   TYPE string.
DATA lv_http   TYPE string.
DATA lv_width  TYPE i.
```

## abapGit Installation for Cloud

1. Open abapGit in your cloud system.
2. Clone `https://github.com/abap2xlsx/abap2xlsx.git`.
3. Map the root package to a cloud-enabled package (e.g. `ZABAP2XLSX`).
4. Map `src/not_cloud/` to a separate package marked as not-for-cloud, or skip activation of those objects.
5. Activate all objects in the main package.
6. Run the demo programs (excluding ALV demos) to verify.

## Known Limitations in Cloud

- **ALV integration** (`ZCL_EXCEL_ALV_*`) is in `not_cloud/` and unavailable in cloud deployments. Use `bind_table()` or cell-by-cell population instead.
- **GUI-based file upload/download** helpers are not available. Use BTP services or ABAP Cloud file APIs.

## Related Resources

- [abap2xlsx `src/not_cloud/` on GitHub](https://github.com/abap2xlsx/abap2xlsx/tree/main/src/not_cloud)
- [SAP Help - ABAP Cloud programming model restrictions](https://help.sap.com/docs/abap-cloud)
- **[ALV Integration](/guide/alv-integration)** - On-premise only
- **[Data Conversion](/guide/data-conversion)** - Cloud-safe patterns
- **[Performance](/guide/performance)** - Applies to all targets
- **[Changelog](/guide/changelog)** - Full history of cloud compatibility commits
