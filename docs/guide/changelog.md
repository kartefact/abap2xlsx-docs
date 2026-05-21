# Changelog

This page summarises significant changes to abap2xlsx, grouped by the month merged into [abap2xlsx/abap2xlsx](https://github.com/abap2xlsx/abap2xlsx) `main`.

For the full commit history see the [GitHub commits list](https://github.com/abap2xlsx/abap2xlsx/commits/main).

---

## 2026

### April 2026

- **npm/tooling dependency updates** (PR [#1335](https://github.com/abap2xlsx/abap2xlsx/pull/1335)) - Internal build tooling updated. No impact on ABAP library consumers.

### February 2026

- **Reader: `totalsRowFunction` preserved on table columns** (PR [#1296](https://github.com/abap2xlsx/abap2xlsx/pull/1296)) - When reading an `.xlsx` file containing a structured table with a totals row, the reader now correctly populates `totalsRowFunction` on each `zcl_excel_table_column`. Previously this attribute was lost on read, breaking round-trip fidelity for SUM/COUNT/AVERAGE totals columns.

### January 2026

- **Cloud: replace `DESCRIBE TABLE LINES`** (PR [#1340](https://github.com/abap2xlsx/abap2xlsx/pull/1340)) - All remaining uses of `DESCRIBE TABLE ... LINES lv_n` replaced with the cloud-compatible `lines( )` built-in. See [Cloud Compatibility](/guide/cloud-compatibility).

---

## 2025

### November 2025

- **Reader: SAP Note 2922674 - XML namespace fix** (PR [#1349](https://github.com/abap2xlsx/abap2xlsx/pull/1349)) - The XML namespace handling from SAP Note 2922674 was previously only applied in the writer. The reader now also handles files containing this namespace declaration, preventing data loss on round-trips.

- **Writer: predefined worksheet node ordering fix** (PR [#1351](https://github.com/abap2xlsx/abap2xlsx/pull/1351)) - `add_ignored_errors` call moved to satisfy the required XML child node order in the worksheet part, eliminating openxml validation warnings on certain Excel versions.

### October 2025

- **Automated tests documentation** (PR [#1342](https://github.com/abap2xlsx/abap2xlsx/pull/1342)) - New `automated-tests.md` added to the source repo with expanded notes on running the demo checker. `CONTRIBUTING.md` also updated.

- **Syntax fix for older systems** (PR [#1344](https://github.com/abap2xlsx/abap2xlsx/pull/1344)) - A syntax expression only accepted by newer kernels corrected for backward compatibility.

### September 2025

- **Cloud: remove `STRTABLE`/`SSTRTABLE`** (PR [#1336](https://github.com/abap2xlsx/abap2xlsx/pull/1336)) - See [Cloud Compatibility](/guide/cloud-compatibility).

- **Cloud: remove `SCRTEXT`** (PR [#1337](https://github.com/abap2xlsx/abap2xlsx/pull/1337)) - See [Cloud Compatibility](/guide/cloud-compatibility).

### August 2025

- **Cloud: remove `SEOCLSNAME` references** (PR [#1331](https://github.com/abap2xlsx/abap2xlsx/pull/1331)) - See [Cloud Compatibility](/guide/cloud-compatibility).

### July 2025

- **Internal refactor: `render_xml_document` constants** (PR [#1327](https://github.com/abap2xlsx/abap2xlsx/pull/1327)) - No user-facing impact.

- **Writer: remove redundant `append_child`** (PR [#1325](https://github.com/abap2xlsx/abap2xlsx/pull/1325)) - `create_simple_element` already appends the node to its parent; the redundant call removed.

### June 2025

- **Worksheet copy constructor improvement** (PR [#1317](https://github.com/abap2xlsx/abap2xlsx/pull/1317)) - Copy logic moved into the copy constructor of `zcl_excel_worksheet`. The comments collection instance is now passed in explicitly, so copying a worksheet correctly carries its comments.

- **RTF loop bug fix in reader** (PR [#1319](https://github.com/abap2xlsx/abap2xlsx/pull/1319)) - `lt_rtf` was not cleared between loop iterations in `zcl_excel_reader_2007->load_worksheet`, causing rich-text content from a previous cell to bleed into subsequent cells.

- **`get_style_from_guid` made public** (PR [#1315](https://github.com/abap2xlsx/abap2xlsx/pull/1315)) - `zcl_excel->get_style_from_guid()` is now public, removing code duplication in `check_rtf`. A comparison operator bug (`>` instead of `<`) also fixed; unit tests added.

- **Comment box parameters consolidated into `ms_box` structure** (PR [#1316](https://github.com/abap2xlsx/abap2xlsx/pull/1316)) - The eight comment box geometry attributes are now wrapped in `ms_box` of type `zcl_excel_comment=>ty_box`. A structured constant `mc_box_default` provides defaults. See [Worksheets](/guide/worksheets).

- **`zcl_excel_converter` LOOP_NORMAL fix** (PR [#1310](https://github.com/abap2xlsx/abap2xlsx/pull/1310)) - `IS ASSIGNED` used incorrectly after a dynamic `ASSIGN`; corrected to `sy-subrc = 0`. Also handles two invalid column names simultaneously. See [Data Conversion](/guide/data-conversion).

- **Cloud: replace `SY-LANGU`/`LANG`** (PR [#1313](https://github.com/abap2xlsx/abap2xlsx/pull/1313)) - See [Cloud Compatibility](/guide/cloud-compatibility).

- **Cloud: replace `IHTTPNVP`** (PR [#1312](https://github.com/abap2xlsx/abap2xlsx/pull/1312)) - See [Cloud Compatibility](/guide/cloud-compatibility).

### May 2025

- **`set_cell` now accepts `XSTRING` values** (PR [#1306](https://github.com/abap2xlsx/abap2xlsx/pull/1306)) - `zcl_excel_worksheet->set_cell()` extended to handle `XSTRING`-typed input directly. See [Basic Usage](/guide/basic-usage).

- **Cloud: replace `TDCWIDTHS`** (PR [#1311](https://github.com/abap2xlsx/abap2xlsx/pull/1311)) - `TDCWIDTHS` type replaced with standard integer `i`. See [Cloud Compatibility](/guide/cloud-compatibility).

### February-March 2025

- **`zcl_excel_converter` performance - pass-by-reference** (PR [#1037](https://github.com/abap2xlsx/abap2xlsx/pull/1037), PR [#1039](https://github.com/abap2xlsx/abap2xlsx/pull/1039)) - Large parameters previously passed by value are now passed by reference. Transparent to callers; reduces memory overhead for large datasets. See [Data Conversion](/guide/data-conversion).

- **Security policy published** (PR [#1289](https://github.com/abap2xlsx/abap2xlsx/pull/1289)) - `SECURITY.md` added establishing the process for privately reporting vulnerabilities. See [Security Policy](/contributing/security).

### January 2025

- **CSV writer: skip hidden rows and columns** (PR [#1268](https://github.com/abap2xlsx/abap2xlsx/pull/1268)) - `zcl_excel_writer_csv` gained `set_skip_hidden_rows()` and `set_skip_hidden_columns()` options. A date-format detection bug for NW 7.52+ systems also fixed. See [CSV Export](/guide/csv-export).

---

## 2024

### September 2024

- **`zcl_excel_converter` autofilter fix** (PR [#1239](https://github.com/abap2xlsx/abap2xlsx/pull/1239)) - Fixed a bug where using the converter with an ALV field catalog including autofilter produced an incorrect or missing filter range. See [Data Conversion](/guide/data-conversion).

---

## How to Read This Changelog

- **PR [#NNNN]** links go directly to the GitHub pull request.
- **Cloud:** entries relate to S/4HANA Cloud / ABAP Cloud compatibility - see [Cloud Compatibility](/guide/cloud-compatibility).
- **Reader:** / **Writer:** refer to `zcl_excel_reader_2007` / `zcl_excel_writer_2007` respectively.
