# Cloud Compatibility (S/4HANA Cloud & BTP ABAP)

This page documents the ongoing effort to make abap2xlsx fully compatible with **SAP S/4HANA Cloud** (ABAP Environment / Steampunk) and **SAP BTP ABAP Environment**, where classic on-premise ABAP syntax is restricted.

## What Is the Cloud Compatibility Track?

Abap2xlsx has historically relied on a small number of classic ABAP language features and built-in variables that are not available in the ABAP Cloud programming model (ABAP for Cloud Development). The maintainers have been systematically removing these dependencies so that the library can be consumed in cloud-native ABAP projects without modification.

Cloud-restricted constructs that have been removed or replaced:

| Removed construct | Replaced with | PR | Merged |
|---|---|---|---|
| `DESCRIBE TABLE … LINES` | `lines( )` built-in function | [#1340](https://github.com/abap2xlsx/abap2xlsx/pull/1340) | Jan 2026 |
| `STRTABLE` / `SSTRTABLE` type references | Standard string/table types | [#1336](https://github.com/abap2xlsx/abap2xlsx/pull/1336) | Sep 2025 |
| `SCRTEXT_S/M/L` (screen-text types) | `string` | [#1337](https://github.com/abap2xlsx/abap2xlsx/pull/1337) | Sep 2025 |
| `SY-LANGU` / `LANG` system field | `cl_abap_context_info=>get_system_language( )` | [#1313](https://github.com/abap2xlsx/abap2xlsx/pull/1313) | Jun 2025 |
| `SEOCLSNAME` type references | `string` | [#1331](https://github.com/abap2xlsx/abap2xlsx/pull/1331) | Aug 2025 |

## Deployment Notes

### Classic On-Premise (ECC, S/4HANA On-Premise)

No changes required on your side. All existing code and transports remain fully supported. The replacements listed above are backward-compatible with the relevant on-premise ABAP releases.

### S/4HANA Cloud / BTP ABAP Environment

Abap2xlsx provides a dedicated `src/not_cloud/` sub-package. Objects inside this folder depend on constructs that are **not** permitted in the ABAP Cloud programming model and must be excluded when deploying to a cloud tenant.

```
abap2xlsx/
  src/
    zcl_excel.clas.abap            ← cloud-safe
    zcl_excel_common.clas.abap     ← cloud-safe
    ...
    not_cloud/                     ← EXCLUDE from cloud deployments
      zcl_excel_*_alv*.clas.abap
      ...
```

When installing via **abapGit** into a cloud system, map the `not_cloud` package to a separate transport and do not activate those objects, or exclude them from the import entirely.

## Checking Your Custom Code

If you extend abap2xlsx with custom code intended for cloud deployment, avoid the following patterns:

```abap
" ❌ Not allowed in ABAP Cloud
DESCRIBE TABLE lt_data LINES lv_count.    " use lines( lt_data ) instead
DATA lv_lang TYPE LANG.                   " use string or CL_ABAP_CONTEXT_INFO
DATA lv_name TYPE SEOCLSNAME.            " use string
DATA lv_text TYPE SCRTEXT_S.             " use string

" ✅ Cloud-compatible equivalents
DATA(lv_count) = lines( lt_data ).
DATA(lv_lang)  = cl_abap_context_info=>get_system_language( ).
DATA lv_name   TYPE string.
DATA lv_text   TYPE string.
```

## abapGit Installation for Cloud

Installing abap2xlsx in a cloud ABAP environment follows the standard abapGit workflow, with the `not_cloud` sub-package excluded:

1. Open abapGit in your cloud system (BTP ABAP Environment or S/4HANA Cloud).
2. Clone `https://github.com/abap2xlsx/abap2xlsx.git`.
3. Map the root package to a cloud-enabled package (e.g., `ZABAP2XLSX`).
4. Map `src/not_cloud/` to a **separate package** marked as not for cloud — or skip activation of those objects entirely.
5. Activate all objects in the main package.
6. Run the demo programs (excluding ALV demos) to verify the installation.

## Known Limitations in Cloud

- **ALV integration** (`ZCL_EXCEL_ALV_*`) is contained within `not_cloud/` and is not available in cloud deployments. Use `bind_table()` or cell-by-cell population instead — see [ALV Integration](/guide/alv-integration) for the on-premise approach and [Data Conversion](/guide/data-conversion) for cloud alternatives.
- **GUI-based file upload/download** helpers are not available. Use BTP services or the ABAP Cloud file APIs provided by your application framework.

## Related Resources

- [abap2xlsx `src/not_cloud/` on GitHub](https://github.com/abap2xlsx/abap2xlsx/tree/main/src/not_cloud)
- [SAP Help — ABAP Cloud programming model restrictions](https://help.sap.com/docs/abap-cloud)
- **[ALV Integration](/guide/alv-integration)** — On-premise only
- **[Data Conversion](/guide/data-conversion)** — Cloud-safe table-to-Excel patterns
- **[Performance Optimization](/guide/performance)** — Applies to all deployment targets
- **[Changelog](/guide/changelog)** — Full history of cloud compatibility commits
