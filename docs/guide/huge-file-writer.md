# Writing Large Files (Streaming Writer)

`zcl_excel_writer_huge_file` is a **streaming XSLT-based writer** designed for workbooks
that are too large to build entirely in memory with `zcl_excel_writer_2007`. It generates
the XML directly to a string or file stream, bypassing the in-memory cell table, and is
therefore suitable for hundreds of thousands of rows.

## When to Use

| Scenario | Recommended writer |
|---|---|
| < ~50 000 rows, full formatting, charts, images | `zcl_excel_writer_2007` |
| 50 000 – 1 000 000+ rows, minimal formatting | `zcl_excel_writer_huge_file` |
| CSV export | `zcl_excel_writer_csv` |

> **Trade-offs:** `zcl_excel_writer_huge_file` supports cell values, basic number formats,
> and simple styles. It does **not** support charts, images, drawings, named ranges, or
> complex merged-cell layouts. Use it only for high-volume tabular data exports.

## Basic Usage

```abap
DATA: lo_excel  TYPE REF TO zcl_excel,
      lo_ws     TYPE REF TO zcl_excel_worksheet,
      lo_writer TYPE REF TO zcl_excel_writer_huge_file.

CREATE OBJECT lo_excel.
lo_ws = lo_excel->get_active_worksheet( ).
lo_ws->set_title( 'Data Export' ).

" Write cells as normal — the streaming writer collects cell data per-row
DATA(lv_row) = 1.
lo_ws->set_cell( ip_column = 'A' ip_row = lv_row ip_value = 'OrderNo' ).
lo_ws->set_cell( ip_column = 'B' ip_row = lv_row ip_value = 'Customer' ).
lo_ws->set_cell( ip_column = 'C' ip_row = lv_row ip_value = 'Amount' ).

LOOP AT lt_orders INTO DATA(ls_order).
  ADD 1 TO lv_row.
  lo_ws->set_cell( ip_column = 'A' ip_row = lv_row ip_value = ls_order-order_no ).
  lo_ws->set_cell( ip_column = 'B' ip_row = lv_row ip_value = ls_order-customer ).
  lo_ws->set_cell( ip_column = 'C' ip_row = lv_row ip_value = ls_order-amount ).
ENDLOOP.

" Use the huge-file writer instead of zcl_excel_writer_2007
CREATE OBJECT lo_writer.
DATA(lv_file) = lo_writer->write_file( lo_excel ).
```

The return type of `write_file` is `xstring`, identical to `zcl_excel_writer_2007`.

## Using `bind_table` with the Streaming Writer

The most efficient pattern for large datasets is `bind_table` followed by the huge-file
writer. `bind_table` populates the in-memory cell table row-by-row from an internal table:

```abap
" Populate worksheet from an internal table
lo_ws->bind_table( ip_table = lt_large_data ).

" Stream it out
CREATE OBJECT lo_writer.
DATA(lv_xstring) = lo_writer->write_file( lo_excel ).
```

## Memory-Saving Tips

1. **Process data in chunks.** If the source data is read from the database, use `PACKAGE SIZE`
   to read and write in batches, clearing the previous batch:
   ```abap
   SELECT * FROM zsales_data INTO TABLE @DATA(lt_chunk) PACKAGE SIZE 10000.
     lo_ws->bind_table( ip_table = lt_chunk ip_start_row = lv_next_row ).
     lv_next_row = lv_next_row + lines( lt_chunk ).
   ENDSELECT.
   ```

2. **Avoid complex styles in loops.** Every distinct style creates a new entry in the
   workbook's style table. Create styles once before the loop and reuse the reference.

3. **Use `bind_table` over `set_cell` loops** where possible — `bind_table` is significantly
   faster for bulk population.

## Limitations

- No charts, images, or drawings
- No named ranges or defined names
- No complex merged cells (header merges before the data are OK)
- No OLE automation features (use `zcl_excel_ole` for that — classic ABAP only)
- Thread safety: the writer is stateless; one instance per `write_file` call is safe

## Comparing the Three Writers

| Feature | `writer_2007` | `writer_huge_file` | `writer_csv` |
|---|:---:|:---:|:---:|
| Full cell styles | ✅ | ⚠️ Basic | ❌ |
| Charts / images | ✅ | ❌ | ❌ |
| Named ranges | ✅ | ❌ | ❌ |
| Page setup / print | ✅ | ⚠️ Limited | ❌ |
| AutoFilter | ✅ | ❌ | via `is_row_hidden` |
| Structured tables | ✅ | ❌ | ❌ |
| Row limit (practical) | ~50 000 | 1 000 000+ | Unlimited |
| Output format | `.xlsx` | `.xlsx` | `.csv` |
| Cloud-compatible | ✅ | ✅ | ✅ |

## Next Steps

- **[Data Conversion](/guide/data-conversion)** — `bind_table` reference
- **[CSV Export](/guide/csv-export)** — for text-only exports
- **[Cloud Compatibility](/guide/cloud-compatibility)** — verify your system supports streaming
