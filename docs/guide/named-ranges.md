# Named Ranges

Named ranges give a human-readable label to a cell or range reference, making formulas self-documenting (`=SUM(SalesQ1)` instead of `=SUM(Sheet1!$B$2:$B$13)`) and enabling dynamic dropdown lists in [Data Validation](./data-validation.md) or lookup anchors in [Template Filling](./template-filling.md).

abap2xlsx manages named ranges through `zcl_excel_range` (one range definition) and `zcl_excel_ranges` (the workbook-level collection at `zcl_excel->ranges`).

## Creating a Named Range

```abap
DATA lo_range TYPE REF TO zcl_excel_range.

lo_range = NEW zcl_excel_range( ).
lo_range->name      = 'SalesData'.
lo_range->value     = 'Sheet1!$B$2:$B$50'.    " absolute reference

lo_excel->ranges->add( io_range = lo_range ).
```

`value` must be a valid Excel range reference. Workbook-scoped names use the sheet name prefix (`SheetName!$col$row`). Sheet-scoped (local) names include the worksheet index in the OOXML `localSheetId` attribute — see below.

## `zcl_excel_range` Attributes

| Attribute | Type | Description |
|---|---|---|
| `name` | `zexcel_name` | The range label (e.g. `SalesData`, `Colours`) |
| `value` | `zexcel_name` | Cell reference or range (e.g. `Sheet1!$A$1:$A$10`) |

Both attributes are simple public data — there are no setter methods.

## Workbook vs. Sheet Scope

By default ranges added via `lo_excel->ranges->add()` are workbook-scoped (visible from any sheet). To create a sheet-scoped (local) name, the OOXML requires a `localSheetId` attribute on the `<definedName>` element. abap2xlsx does not expose a separate API for this — use a workbook-scoped name with an explicit sheet prefix in `value` for the same practical effect.

## Iterating Existing Ranges

```abap
DATA lo_iter  TYPE REF TO zcl_excel_collection_iterator.
DATA lo_obj   TYPE REF TO object.
DATA lo_range TYPE REF TO zcl_excel_range.

lo_iter = lo_excel->ranges->get_iterator( ).
WHILE lo_iter->has_next( ) = abap_true.
  lo_obj   = lo_iter->get_next( ).
  lo_range ?= lo_obj.
  WRITE: / lo_range->name, lo_range->value.
ENDWHILE.
```

## Common Use Cases

### Data Validation Dropdown from a Named Range

Define the list once as a named range, then reference it in multiple validation rules:

```abap
" 1. Create the named range
DATA lo_range TYPE REF TO zcl_excel_range.
lo_range = NEW zcl_excel_range( ).
lo_range->name  = 'ColourList'.
lo_range->value = 'Lists!$A$2:$A$6'.   " "Lists" sheet, rows 2-6
lo_excel->ranges->add( io_range = lo_range ).

" 2. Reference it in data validation
DATA lo_val TYPE REF TO zcl_excel_data_validation.
lo_val = NEW zcl_excel_data_validation( ).
lo_val->type        = zcl_excel_data_validation=>c_type_list.
lo_val->formula1    = 'ColourList'.    " named range — no quotes!
lo_val->cell_column = 'D'.
lo_val->cell_row    = 2.
lo_worksheet->data_validations->add( lo_val ).
```

### Print Area

A print area is a workbook-level named range with the reserved name `Print_Area`:

```abap
lo_range = NEW zcl_excel_range( ).
lo_range->name  = 'Print_Area'.
lo_range->value = 'Sheet1!$A$1:$H$40'.
lo_excel->ranges->add( io_range = lo_range ).
```

::: tip Print area via worksheet
For a simpler API see `zcl_excel_worksheet->set_print_area()` which handles the named range creation internally.
:::

### Dynamic Formula Reference

```abap
" Sum of named range — cleaner than cell-address arithmetic
lo_worksheet->set_cell(
  ip_column = 'J'
  ip_row    = 1
  ip_value  = '=SUM(SalesData)'
  ip_formula = abap_true
).
```

## Reading Named Ranges from an Existing File

`zcl_excel_reader_2007` populates `lo_excel->ranges` from the `<definedNames>` section of `workbook.xml` when it reads an xlsx file. All workbook-scoped ranges (including auto-filter ranges and print areas written by Excel) will be present after reading.

## Named Ranges and Template Filling

`zcl_excel_fill_template` uses named ranges extensively as its primary mechanism for locating template regions. See [Template Filling](./template-filling.md) for the full workflow.

## Limitations

- Names containing spaces must be enclosed in single quotes when used inside Excel formulas: `'Sales Data'`.
- Excel reserved names (`Print_Area`, `Print_Titles`, `_xlnm.FilterDatabase`) are written as-is — no special handling.
- Dynamic array spill ranges (`Sheet1!$A$1#`) introduced in Excel 365 are not supported.
