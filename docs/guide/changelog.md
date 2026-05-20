# Changelog

This page summarises significant changes to abap2xlsx. Entries are grouped by the month they were merged into the `main` branch of [abap2xlsx/abap2xlsx](https://github.com/abap2xlsx/abap2xlsx).

For the full commit history, see the [GitHub commits list](https://github.com/abap2xlsx/abap2xlsx/commits/main).

---

## 2026

### April 2026

- **npm/tooling dependency updates** (PR [#1335](https://github.com/abap2xlsx/abap2xlsx/pull/1335)) — Internal build tooling updated. No impact on ABAP library consumers.

### February 2026

- **Reader: `totalsRowFunction` preserved on table columns** (PR [#1296](https://github.com/abap2xlsx/abap2xlsx/pull/1296)) — When reading an `.xlsx` file that contains a structured table with a totals row, the reader now correctly populates `totalsRowFunction` on each `zcl_excel_table_column`. Previously this attribute was lost on read, breaking round-trip fidelity for files with SUM/COUNT/AVERAGE totals columns. See [Reading Excel Files — Reading Table Column Totals Row Functions](/guide/reading-excel#reading-table-column-totals-row-functions).

### January 2026

- **Cloud: replace `DESCRIBE TABLE LINES`** (PR [#1340](https://github.com/abap2xlsx/abap2xlsx/pull/1340)) — All remaining uses of `DESCRIBE TABLE … LINES lv_n` replaced with the cloud-compatible `lines( )` built-in. See [Cloud Compatibility](/guide/cloud-compatibility).

---

## 2025

### November 2025

- **Reader: SAP Note 2922674 — XML namespace fix** (PR [#1349](https://github.com/abap2xlsx/abap2xlsx/pull/1349)) — The XML namespace handling introduced by SAP Note 2922674 was previously only applied in the writer. The reader now also handles files that contain this namespace declaration, preventing data loss when round-tripping files produced on certain SAP releases. See [Reading Excel Files — SAP Note 2922674](/guide/reading-excel#sap-note-2922674--xml-namespace-handling).

- **Writer: predefined worksheet node ordering fix** (PR [#1351](https://github.com/abap2xlsx/abap2xlsx/pull/1351)) — The call to `add_ignored_errors` in `zcl_excel_writer_2007` was moved to satisfy the required order of XML child nodes in the worksheet part. Files that previously produced an `openxml` validation warning in certain Excel versions are now written with a fully conformant node sequence.

### October 2025

- **Automated tests documentation** (PR [#1342](https://github.com/abap2xlsx/abap2xlsx/pull/1342)) — A new `automated-tests.md` was added to the source repository with expanded notes on running the demo checker, including tips for older SAP releases. `CONTRIBUTING.md` was also updated.

- **Syntax fix for older systems** (PR [#1344](https://github.com/abap2xlsx/abap2xlsx/pull/1344)) — A syntax expression written in a form only accepted by newer kernels was corrected to ensure backward compatibility with older on-premise systems.

### September 2025

- **Cloud: remove `STRTABLE`/`SSTRTABLE`** (PR [#1336](https://github.com/abap2xlsx/abap2xlsx/pull/1336)) — References to the non-cloud-compatible `STRTABLE` and `SSTRTABLE` built-in types replaced with standard equivalents.

- **Cloud: remove `SCRTEXT`** (PR [#1337](https://github.com/abap2xlsx/abap2xlsx/pull/1337)) — References to `SCRTEXT_S/M/L` types replaced with `string`.

### August 2025

- **Cloud: remove `SEOCLSNAME` references** (PR [#1331](https://github.com/abap2xlsx/abap2xlsx/pull/1331)) — The `SEOCLSNAME` type replaced with `string` in all relevant declarations.

### July 2025

- **Internal refactor: `render_xml_document` constants** (PR [#1327](https://github.com/abap2xlsx/abap2xlsx/pull/1327)) — Internal CONSTANTS replaced with literals in the writer method as a preparatory refactor. No user-facing impact.

- **Writer: remove redundant `append_child`** (PR [#1325](https://github.com/abap2xlsx/abap2xlsx/pull/1325)) — `create_simple_element` already appends the new node to its parent; the redundant subsequent `append_child` call was removed. Partial fix for [#1324](https://github.com/abap2xlsx/abap2xlsx/issues/1324).

### June 2025

- **Worksheet copy constructor improvement** (PR [#1317](https://github.com/abap2xlsx/abap2xlsx/pull/1317)) — Copy logic moved into the copy constructor of `zcl_excel_worksheet`. The comments collection instance is now passed in explicitly, so copying a worksheet correctly carries its comments. `get_comments()` copies the collection by default. See [Worksheets — Reading Comments](/guide/worksheets#reading-comments-copy-semantics).

- **RTF loop bug fix in reader** (PR [#1319](https://github.com/abap2xlsx/abap2xlsx/pull/1319)) — `lt_rtf` was not cleared between loop iterations in `zcl_excel_reader_2007->load_worksheet`, causing rich-text content from a previous cell to bleed into subsequent cells.

- **`get_style_from_guid` made public** (PR [#1315](https://github.com/abap2xlsx/abap2xlsx/pull/1315)) — `zcl_excel->get_style_from_guid()` is now public. Code duplication in `zcl_excel_worksheet->check_rtf` removed. A comparison operator bug (`>` instead of `<`) in `check_rtf` also fixed. Unit tests added. See [Reading Excel Files — Accessing `get_style_from_guid` Publicly](/guide/reading-excel#accessing-get_style_from_guid-publicly).

- **Comment box parameters consolidated into `ms_box` structure** (PR [#1316](https://github.com/abap2xlsx/abap2xlsx/pull/1316)) — The eight attributes controlling comment box geometry (`bottom_offset`, `bottom_row`, `left_column`, `left_offset`, `right_column`, `right_offset`, `top_offset`, `top_row`) are now wrapped in a structure `ms_box` of type `zcl_excel_comment=>ty_box`. A structured constant `mc_box_default` provides default values. See [Worksheets — Comment Box Positioning](/guide/worksheets#comment-box-positioning--updated-api-2025-06).

- **`zcl_excel_converter` error handling fix** (PR [#1310](https://github.com/abap2xlsx/abap2xlsx/pull/1310)) — Fixed a bug in `LOOP_NORMAL` where `IS ASSIGNED` was used incorrectly after a dynamic `ASSIGN`; corrected to check `SY-SUBRC = 0`. Also handles the case of two invalid column names simultaneously.

- **Cloud: replace `SY-LANGU`/`LANG`** (PR [#1313](https://github.com/abap2xlsx/abap2xlsx/pull/1313)) — References to the `LANG` type and `SY-LANGU` system field replaced with `cl_abap_context_info=>get_system_language( )`. See [Cloud Compatibility](/guide/cloud-compatibility).

- **Fix broken CONTRIBUTING.md links** (PR [#1320](https://github.com/abap2xlsx/abap2xlsx/pull/1320)) — Links to the SAP Community platform updated following the 2024 platform migration.

---

## How to Read This Changelog

- **PR [#NNNN]** links go directly to the GitHub pull request where you can read the full diff and discussion.
- Entries marked **Cloud:** relate to S/4HANA Cloud / ABAP Cloud programming model compatibility — see [Cloud Compatibility](/guide/cloud-compatibility).
- Entries marked **Reader:** / **Writer:** refer to `zcl_excel_reader_2007` / `zcl_excel_writer_2007` respectively.
