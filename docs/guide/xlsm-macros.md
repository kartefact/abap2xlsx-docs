# Macro-Enabled Workbooks (XLSM)

The `.xlsm` format is an Office Open XML workbook that can contain **VBA macros**. abap2xlsx
provides a dedicated reader (`zcl_excel_reader_xlsm`) and writer (`zcl_excel_writer_xlsm`)
for macro-enabled workbooks.

> ⚠️ **Important:** abap2xlsx does **not** create or edit VBA code. Its XLSM support is
> limited to **preserving** the existing VBA project stored in `vbaProject.bin` inside the
> package. Use the XLSM reader to load a pre-authored macro workbook, modify the data with
> the normal abap2xlsx API, then save it back with the XLSM writer — the macros are
> round-tripped unchanged.

## Reading an XLSM File

```abap
DATA: lo_reader    TYPE REF TO zcl_excel_reader_xlsm,
      lo_excel     TYPE REF TO zcl_excel.

" Load the .xlsm binary — source can be BDS, MIME, GUI_UPLOAD, etc.
DATA(lv_xlsm_xstring) = load_xlsm_from_mime( ).  " your helper

CREATE OBJECT lo_reader.
lo_excel = lo_reader->load( lv_xlsm_xstring ).

" All standard worksheet operations work on the loaded workbook
DATA(lo_ws) = lo_excel->get_active_worksheet( ).
DATA(lv_value) = lo_ws->get_cell( ip_column = 'A' ip_row = 1 )-value.
```

`zcl_excel_reader_xlsm` inherits from `zcl_excel_reader_2007`. It overrides only the parts
that differ for `.xlsm`: it reads the `vbaProject.bin` part and stores it internally so the
writer can re-embed it.

## Writing an XLSM File

```abap
DATA: lo_writer TYPE REF TO zcl_excel_writer_xlsm.

" Modify worksheet data as normal, then write back as .xlsm
lo_ws->set_cell( ip_column = 'B' ip_row = 2 ip_value = 'Updated by ABAP' ).

CREATE OBJECT lo_writer.
DATA(lv_xlsm_out) = lo_writer->write_file( lo_excel ).

" lv_xlsm_out is an XSTRING ready for download / BDS storage
```

`zcl_excel_writer_xlsm` inherits from `zcl_excel_writer_2007`. It re-embeds the
`vbaProject.bin` part and adjusts the `[Content_Types].xml` and `.rels` entries to
declare the file as a macro-enabled workbook.

## Typical Workflow: XLSM Template Fill

A common pattern is to maintain an XLSM template with macros in BDS or MIME, fill it at
runtime with ABAP data, and send it to the user. The macros run automatically when the user
opens the file:

```abap
METHOD get_macro_report.
  " 1. Load the pre-authored .xlsm template from MIME
  DATA(lv_template) = load_from_mime( 'ZMACRO_REPORT_TEMPLATE' ).

  " 2. Read with the XLSM reader — preserves vbaProject.bin
  DATA(lo_reader) = NEW zcl_excel_reader_xlsm( ).
  DATA(lo_excel)  = lo_reader->load( lv_template ).

  " 3. Fill data using the normal API or zcl_excel_fill_template
  DATA(lo_ws) = lo_excel->get_worksheet_by_name( 'Data' ).
  lo_ws->bind_table( ip_table = mt_report_data ).

  " 4. Write back as .xlsm — macros intact
  DATA(lo_writer) = NEW zcl_excel_writer_xlsm( ).
  DATA(lv_output) = lo_writer->write_file( lo_excel ).

  " 5. Deliver to front-end
  download_file(
    iv_data     = lv_output
    iv_filename = 'report.xlsm'
    iv_mimetype = 'application/vnd.ms-excel.sheet.macroEnabled.12'
  ).
ENDMETHOD.
```

## Limitations

- abap2xlsx cannot **create** VBA code or modify existing macros programmatically.
- The `vbaProject.bin` binary is copied verbatim — no inspection or modification is
  possible from ABAP.
- Macros that reference specific sheet names will break if you rename sheets.
- When adding new worksheets and then writing as XLSM, the new sheets are data-only;
  macros that dynamically reference sheet counts or names may behave differently.
- If the source file is an `.xlsx` (no `vbaProject.bin`), writing it with
  `zcl_excel_writer_xlsm` produces a valid `.xlsm` but without any macros — functionally
  identical to an `.xlsx`.

## Next Steps

- **[Template Filling](/guide/template-filling)** — fill an XLSM template with named-range data binding
- **[Reading Excel](/guide/reading-excel)** — general reading guide
- **[Cloud Compatibility](/guide/cloud-compatibility)** — XLSM is fully cloud-compatible
