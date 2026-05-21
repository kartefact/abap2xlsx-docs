# Suppressing Cell Validation Errors

Excel marks certain cells with a green triangle warning when it detects a potential issue — numbers stored as text, formulas that differ from adjacent cells, and so on. `zcl_excel_worksheet` lets you programmatically suppress any of these warnings for specific cell ranges using `set_ignored_errors` / `get_ignored_errors`.

## How it works

Each worksheet holds a hashed table of type `mty_th_ignored_errors`, keyed on `cell_coords`. Each entry targets one cell address or range and carries ten boolean flags, one per Excel warning category. You build the table in ABAP and hand it to the worksheet; the writer serialises it into the `<ignoredErrors>` element of the sheet XML.

```abap
DATA: lt_ignored TYPE zcl_excel_worksheet=>mty_th_ignored_errors,
      ls_ignored TYPE zcl_excel_worksheet=>mty_s_ignored_errors.

" Suppress "number stored as text" warning for column A rows 2-100
ls_ignored-cell_coords          = 'A2:A100'.
ls_ignored-number_stored_as_text = abap_true.
INSERT ls_ignored INTO TABLE lt_ignored.

" Suppress formula-differs warning on a totals row
CLEAR ls_ignored.
ls_ignored-cell_coords = 'B101:Z101'.
ls_ignored-formula     = abap_true.
INSERT ls_ignored INTO TABLE lt_ignored.

lo_worksheet->set_ignored_errors( lt_ignored ).
```

## `mty_s_ignored_errors` — flag reference

| Field | Excel warning suppressed |
|---|---|
| `eval_error` | Formula evaluates to an error (e.g. `#DIV/0!`) |
| `two_digit_text_year` | Year represented as two digits in a text-formatted cell |
| `number_stored_as_text` | Number entered or pasted as text, or preceded by an apostrophe |
| `formula` | Formula in a region differs from other formulas in the same region |
| `formula_range` | Formula omits cells in a contiguous region |
| `unlocked_formula` | Unlocked (unprotected) cell contains a formula |
| `empty_cell_reference` | Formula references empty cells |
| `list_data_validation` | Cell value does not comply with a data-validation rule |
| `calculated_column` | Cell in a table column has a formula different from the column formula |

> Only flags set to `abap_true` are written; the others are omitted from the XML.

## Reading back the current state

```abap
DATA(lt_current) = lo_worksheet->get_ignored_errors( ).
```

The returned table is a snapshot — modifying it does not affect the worksheet. Call `set_ignored_errors` again to replace the entire set.

## `cell_coords` format

Accepted formats match the standard Excel reference syntax:

| Format | Example | Meaning |
|---|---|---|
| Single cell | `'C5'` | One cell |
| Range | `'A2:A100'` | Contiguous block |
| Multi-cell list | `'A1 B3 C7'` | Space-separated addresses |

## Typical use cases

### IDoc / interface data imports
When you populate cells from parsed strings (e.g. IDoc segment fields), numeric fields may arrive as character strings. Rather than converting every field, suppress the warning for the data columns:

```abap
ls_ignored-cell_coords          = |B2:B{ lv_last_row }|.
ls_ignored-number_stored_as_text = abap_true.
```

### ALV-driven reports with formula total rows
When `bind_table` writes a totals row whose formulas differ slightly from the body column formulas:

```abap
ls_ignored-cell_coords = |A{ lv_total_row }:{ lv_last_col_alpha }{ lv_total_row }|.
ls_ignored-formula     = abap_true.
```

### Protecting formula cells without locking the sheet
If the workbook is not fully protected but formulas should not generate warnings:

```abap
ls_ignored-cell_coords       = 'C1:Z1000'.
ls_ignored-unlocked_formula  = abap_true.
```

## See also

- [Data Validation](./data-validation.md) — constraining what users can enter
- [Workbook Security](./workbook-security.md) — sheet and workbook protection
- [Worksheets](./worksheets.md) — full worksheet API reference
