# Template Filling

`zcl_excel_fill_template` lets you design an Excel workbook as a **visual template** in SE80 or any
Spreadsheet editor, then fill it at runtime by binding ABAP data to named ranges and
cell-level placeholder variables. This is the right approach when a fixed layout must be
preserved exactly — e.g. print-ready forms, management reports with a fixed header/footer,
or legal documents.

## How It Works

The engine reads all **named ranges** already defined in the workbook. Named ranges that span
only **full rows** (column A through XFD) become **repeating range blocks** — rows that are
stamped once for every entry in a data table. Named ranges that span a partial column range
are silently ignored by the engine (no error, just skipped).

Cell values containing **placeholder tokens** in the form `[FIELDNAME]` are replaced at
runtime with the corresponding field value from the bound ABAP structure or table row.

Nested ranges are fully supported: a named range whose rows fall inside another named range
creates a **parent-child hierarchy** that maps directly to nested ABAP internal tables.

## Quick Start

### 1. Design the Template Workbook

Open Excel (or create the workbook programmatically) and define named ranges over entire rows:

```
Workbook: invoice_template.xlsx
  Sheet "Invoice"
    Row 1   : Fixed header — company name, logo, etc.
    Rows 4–4: Named range  ITEMS     (one row per line item)
    Row 5   : Named range  TOTALS    (summary row repeated once for the summary data)
```

In cell A4 place the text  `[ITEM_NO]`,  cell B4  `[DESCRIPTION]`,  cell C4  `[AMOUNT]`.
The token names map **case-insensitively** to structure field names.

### 2. Define the ABAP Data Structures

Data is passed through `zcl_excel_template_data=>ts_template_data_sheet`, which groups
everything belonging to one sheet:

```abap
TYPES:
  BEGIN OF ts_item,
    item_no     TYPE i,
    description TYPE string,
    amount      TYPE p DECIMALS 2,
  END OF ts_item,
  tt_items TYPE STANDARD TABLE OF ts_item WITH DEFAULT KEY,

  BEGIN OF ts_totals,
    net_amount TYPE p DECIMALS 2,
    tax        TYPE p DECIMALS 2,
    total      TYPE p DECIMALS 2,
  END OF ts_totals,

  " Top-level data structure whose component names match the named ranges
  BEGIN OF ts_invoice_data,
    items   TYPE tt_items,    " Component name = named range ITEMS
    totals  TYPE ts_totals,   " Component name = named range TOTALS
    " Scalar fields here become top-level [TOKENS] outside any range
    invoice_no TYPE string,
    due_date   TYPE d,
  END OF ts_invoice_data.
```

> **Key rule:** every component name in the top-level structure must match a named range
> name **or** a `[TOKEN]` placeholder in the sheet, compared case-insensitively after
> stripping brackets.

### 3. Load the Template and Fill It

```abap
DATA: lo_reader   TYPE REF TO zcl_excel_reader_2007,
      lo_excel    TYPE REF TO zcl_excel,
      lo_filler   TYPE REF TO zcl_excel_fill_template,
      lo_writer   TYPE REF TO zcl_excel_writer_2007.

" a) Load the pre-designed template workbook
DATA(lv_template_xstr) = get_template_as_xstring( ).  " load from MIME, BDS, file, etc.
CREATE OBJECT lo_reader.
lo_excel = lo_reader->load( lv_template_xstr ).

" b) Analyse the template (reads named ranges, discovers variables)
lo_filler = zcl_excel_fill_template=>create( lo_excel ).

" c) Prepare the data
DATA(ls_data) = VALUE ts_invoice_data(
  invoice_no = 'INV-2026-0042'
  due_date   = '20260630'
  items      = VALUE tt_items(
    ( item_no = 1  description = 'Consulting services'  amount = '5000.00' )
    ( item_no = 2  description = 'Travel expenses'      amount = '800.00'  )
  )
  totals     = VALUE ts_totals(
    net_amount = '5800.00'
    tax        = '1044.00'
    total      = '6844.00'
  )
).

" d) Build the sheet-binding descriptor
DATA(ls_sheet_data) = VALUE zcl_excel_template_data=>ts_template_data_sheet(
  sheet = 'Invoice'   " exact tab name — case-sensitive
  data  = REF #( ls_data )
).

" e) Fill the sheet
lo_filler->fill_sheet( ls_sheet_data ).

" f) Serialise
CREATE OBJECT lo_writer.
DATA(lv_output) = lo_writer->write_file( lo_excel ).
```

## Nested Ranges (Repeating Sub-Tables)

When a named range row span sits **entirely within** another named range, the engine treats
it as a child range. The parent structure's component with that name must be an internal
table.

```
Sheet "Report"
  Rows 3–10 : Named range DEPARTMENT   (outer loop — one block per department)
  Rows 5–8  : Named range EMPLOYEE     (inner loop — one block per employee within a dept)
```

```abap
TYPES:
  BEGIN OF ts_employee,
    name   TYPE string,
    salary TYPE p DECIMALS 2,
  END OF ts_employee,
  tt_employees TYPE STANDARD TABLE OF ts_employee WITH DEFAULT KEY,

  BEGIN OF ts_department,
    dept_name  TYPE string,
    employee   TYPE tt_employees,  " component = nested named range EMPLOYEE
  END OF ts_department,
  tt_departments TYPE STANDARD TABLE OF ts_department WITH DEFAULT KEY,

  BEGIN OF ts_report_data,
    department TYPE tt_departments,  " component = outer named range DEPARTMENT
    report_title TYPE string,
  END OF ts_report_data.
```

The engine recursively stamps the inner `EMPLOYEE` block for each row in `tt_employees`,
then advances to the next `DEPARTMENT` block.  Row offsets and merged cells are recalculated
automatically.

## Variable Token Resolution

Tokens are matched by a **REGEX** `\[[^\]]*\]` — anything between square brackets.
Resolution rules:

1. The token name is uppercased and stripped of brackets/spaces.
2. The engine looks for a structure component of that name in the **innermost enclosing range**
   (or in the top-level data structure if the token is outside all ranges).
3. If the cell had a number/date/time format style applied in the template, the engine
   preserves that data type after substitution (numeric, date, time, or text).
4. If multiple tokens appear in a single cell (e.g. `[FIRST_NAME] [LAST_NAME]`), all are
   replaced via string concatenation; the result is always stored as a text cell.
5. If a single token is the **only content** of a cell, the engine adopts the ABAP value's
   native data type — numbers are written as numeric cells, dates as date-formatted cells.

## `create` Factory Method

```abap
" Signature:
CLASS-METHODS create
  IMPORTING io_excel                 TYPE REF TO zcl_excel
  RETURNING VALUE(eo_template_filler) TYPE REF TO zcl_excel_fill_template
  RAISING   zcx_excel.
```

`create` performs four internal steps:
1. **`get_range`** — iterates all worksheets and collects named ranges spanning full rows.
2. **`discard_overlapped`** — removes partially-overlapping ranges (only non-overlapping or
   fully-nested ranges are valid).
3. **`sign_range`** — assigns numeric IDs and builds the parent-child hierarchy.
4. **`find_var`** — scans every cell for `[TOKEN]` patterns and records which range scope
   each token belongs to.

## `fill_sheet` Method

```abap
" Signature:
METHODS fill_sheet
  IMPORTING iv_data TYPE zcl_excel_template_data=>ts_template_data_sheet
  RAISING   zcx_excel.

" ts_template_data_sheet:
"   sheet  TYPE zexcel_sheet_title  — exact tab name
"   data   TYPE REF TO data        — reference to the top-level data structure
```

`fill_sheet` processes one sheet at a time. Call it once per sheet that contains template
content. The method:
1. Makes a copy of the sheet's cell content and merged-cell table.
2. Recursively expands named range blocks.
3. Substitutes `[TOKEN]` placeholders with live values.
4. Writes the result back to the worksheet, preserving all styles and merges.

## Read-Only Diagnostic Attributes

After `create`, three attributes are available for debugging:

| Attribute | Type | Description |
|---|---|---|
| `mt_sheet` | `tt_sheet_titles` | All sheet titles found in the workbook |
| `mt_range` | `tt_ranges` | All valid named ranges discovered (with parent IDs) |
| `mt_var` | `tt_variables` | All `[TOKEN]` variables found, with their parent range IDs |
| `mt_name_styles` | `tt_name_styles` | Per-variable style counters (numeric/date/time/text occurrences) |

## Known Limitations

- Only **row-oriented** named ranges (spanning column A through XFD) are recognised.
  Column-only or partial-row ranges are silently ignored.
- A named range's data component must be an internal table for it to repeat. Scalar
  components at the top level work as simple token replacements.
- Merged cells inside repeating ranges are correctly re-merged after expansion.
- The engine does **not** recalculate formulas that reference the expanded rows by
  absolute row number — use named ranges or table references in formulas instead.
- Images, charts, and drawings attached to rows inside a repeating range are **not** copied
  with the repeated rows.
- `fill_sheet` must be called **after** `create`. Calling it on a different `zcl_excel`
  instance than the one passed to `create` raises `zcx_excel`.

## Next Steps

- **[Reading Excel](/guide/reading-excel)** — load the template from a file or MIME repository
- **[Worksheets](/guide/worksheets)** — named ranges, worksheet management
- **[Data Conversion](/guide/data-conversion)** — `bind_table` as an alternative for simpler
  table-dump scenarios
- **[Changelog](/guide/changelog)** — history of template-filling changes
