# Data Conversion

Abap2xlsx maps ABAP data types to Excel cell types and applies conversion exits where necessary. This page documents how common ABAP types are handled, what conversion helpers are available, and how to control the conversion behaviour.

## ABAP → Excel type mapping

| ABAP type | ABAP type kind | Excel cell type | Notes |
|---|---|---|---|
| `I` Integer | `i` | Number | Stored as integer |
| `F` Float | `f` | Number | IEEE 754 double |
| `P` Packed | `p` | Number | Decimal separator locale-aware |
| `DECFLOAT16` | `a` | Number | 16-digit decimal float |
| `DECFLOAT34` | `e` | Number | 34-digit decimal float |
| `D` Date | `d` | Number (date serial) | Days since 1900-01-00; worksheet date format applied |
| `T` Time | `t` | Number (time fraction) | Fractional day (0.5 = 12:00:00) |
| `UTCLONG` | `p` (internal) | Number (datetime serial) | **S/4HANA only** — see below |
| `C` Character | `c` | String | Leading zeros preserved |
| `N` Numeric text | `n` | String or Number | Treated as string to preserve leading zeros |
| `STRING` | `g` | String | |
| `X` / `XSTRING` | `x` / `y` | Not directly mapped | Encode as base64 string or embed as a drawing |

### UTCLONG support (S/4HANA)

`UTCLONG` is the S/4HANA UTC timestamp type (16-byte packed, microsecond precision). When `set_cell` receives a value whose runtime type resolves to the internal typekind `'p'` and the data reference matches the `variable_utclong` pattern, the value is converted to an Excel datetime serial number (combined date + time fraction).

This conversion is transparent — simply pass the `UTCLONG` field directly:

```abap
DATA: lv_ts TYPE utclong.
GET TIME STAMP FIELD lv_ts.

lo_worksheet->set_cell(
  ip_column = 3
  ip_row    = 5
  ip_value  = lv_ts ).
```

Apply a combined date-time number format to the cell style so Excel renders it correctly:

```abap
DATA(lv_style) = lo_worksheet->change_cell_style(
  ip_column                     = 3
  ip_row                        = 5
  ip_number_format_format_code  = 'YYYY-MM-DD HH:MM:SS' ).
```

> `UTCLONG` is only available on systems running SAP_BASIS 7.55 or higher (S/4HANA 2020+). On ECC / older NetWeaver systems, use `TIMESTAMP` (`P` type, 15 digits) and format it as a date/time string before passing to `set_cell`.

## `ip_abap_type` override

When automatic type detection produces the wrong Excel type, pass `ip_abap_type` explicitly:

```abap
" Force a numeric string field to be treated as a number
lo_worksheet->set_cell(
  ip_column    = 2
  ip_row       = 4
  ip_value     = '000123'
  ip_abap_type = cl_abap_typedescr=>typekind_int ).

" Force a packed field to be treated as text (preserve formatting)
lo_worksheet->set_cell(
  ip_column    = 2
  ip_row       = 5
  ip_value     = lv_packed
  ip_abap_type = cl_abap_typedescr=>typekind_char ).
```

The `ip_abap_type` parameter accepts constants from `cl_abap_typedescr` (`typekind_int`, `typekind_char`, `typekind_float`, etc.).

## `ip_data_type` — explicit Excel type

For direct control of the Excel cell type tag use `ip_data_type` (type `zexcel_cell_data_type`):

| Constant | XML value | When to use |
|---|---|---|
| `zcl_excel_worksheet=>c_cell_type_string` | `s` | Force string shared table entry |
| `zcl_excel_worksheet=>c_cell_type_formula` | `f` | Formula result type override |
| `zcl_excel_worksheet=>c_cell_type_number` | `n` | Numeric (default for numbers) |
| `zcl_excel_worksheet=>c_cell_type_boolean` | `b` | TRUE / FALSE |

## Conversion exits

### `ip_conv_exit_length`

When `abap_true`, applies the `LENGTH` conversion exit to a field before writing. This pads the value to the field's declared length, which preserves leading spaces in fixed-length character fields:

```abap
lo_worksheet->set_cell(
  ip_column           = 1
  ip_row              = 5
  ip_value            = lv_field
  ip_conv_exit_length = abap_true ).
```

The same parameter is available on `bind_table` as `ip_conv_exit_length` (applies to all fields in the table).

### `ip_conv_curr_amt_ext`

When `abap_true` on `bind_table`, applies the external currency-amount conversion exit to all `CURR` fields. This formats the value according to the currency's decimal places definition in table `TCURX`:

```abap
lo_worksheet->bind_table(
  ip_table             = lt_sales
  it_field_catalog     = lt_catalog
  ip_conv_curr_amt_ext = abap_true ).
```

## Currency fields in `bind_table`

For proper currency formatting, populate the `currency` field in `zexcel_s_fieldcatalog` and the corresponding `ref_field` pointing to the currency code column:

```abap
ls_cat-fieldname = 'AMOUNT'.
ls_cat-currency  = 'EUR'.
ls_cat-ref_field = 'WAERS'.   " column holding the currency key
APPEND ls_cat TO lt_catalog.
```

The writer then formats the number cell with the matching Excel currency number format.

## `zcl_excel_common` helper methods

| Method | Purpose |
|---|---|
| `convert_column2alpha` | Integer column number → alpha (`1` → `'A'`, `26` → `'Z'`, `27` → `'AA'`) |
| `convert_column2int` | Alpha column → integer (`'AA'` → `27`) |
| `convert_date2excel` | ABAP `D` date → Excel serial number |
| `convert_time2excel` | ABAP `T` time → fractional day |
| `convert_excel2date` | Excel serial → ABAP date |
| `convert_columnrow2alpha` | Column int + row int → cell reference string (`2,3` → `'B3'`) |
| `excel_string_to_date` | Parse a date string from a cell value |
| `excel_string_to_time` | Parse a time string from a cell value |

```abap
" Column number to letter
DATA(lv_alpha) = zcl_excel_common=>convert_column2alpha( 4 ).  " → 'D'

" ABAP date to Excel serial
DATA(lv_serial) = zcl_excel_common=>convert_date2excel( sy-datum ).
```

## See also

- [Writing Cells](./worksheets.md) — `set_cell` signature and examples
- [Reading Excel Files](./reading-excel.md) — reverse mapping (Excel → ABAP)
- [Convert to Table](./convert-to-table.md) — `convert_to_table` for bulk sheet reads
- [bind_table and Field Catalog](./excel-tables.md)
