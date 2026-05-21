# Data Validation

Data validation lets you constrain what a user can enter into a cell — restricting input to a dropdown list, a numeric range, a date window, or a custom formula. abap2xlsx models validation through `zcl_excel_data_validation` (one rule) and `zcl_excel_data_validations` (the worksheet-level collection).

## Quick Example

The simplest use case: restrict a cell to a dropdown list of values.

```abap
DATA lo_validation TYPE REF TO zcl_excel_data_validation.

lo_validation = NEW zcl_excel_data_validation( ).
lo_validation->type          = zcl_excel_data_validation=>c_type_list.
lo_validation->formula1      = '"Red,Green,Blue"'.  " inline list
lo_validation->cell_column   = 'B'.
lo_validation->cell_row      = 2.
lo_validation->showdropdown  = abap_false.  " abap_false = show the arrow
lo_validation->showinputmessage = abap_true.
lo_validation->prompttitle   = 'Colour'.
lo_validation->prompt        = 'Pick a colour from the list'.

lo_worksheet->data_validations->add( lo_validation ).
```

## Validation Types

Set `zcl_excel_data_validation->type` to one of the constants:

| Constant | Value | Meaning |
|---|---|---|
| `c_type_none` | `none` | No constraint (removes existing) |
| `c_type_list` | `list` | Dropdown list of values |
| `c_type_whole` | `whole` | Integer only |
| `c_type_decimal` | `decimal` | Decimal number |
| `c_type_date` | `date` | Date range |
| `c_type_time` | `time` | Time range |
| `c_type_textlength` | `textLength` | String length constraint |
| `c_type_custom` | `custom` | Custom formula |

## Operators

For numeric, date, time, and text-length types you also set `operator`:

| Constant | Value |
|---|---|
| `c_operator_between` | `between` |
| `c_operator_notbetween` | `notBetween` |
| `c_operator_equal` | `equal` |
| `c_operator_notequal` | `notEqual` |
| `c_operator_greaterthan` | `greaterThan` |
| `c_operator_greaterthanorequal` | `greaterThanOrEqual` |
| `c_operator_lessthan` | `lessThan` |
| `c_operator_lessthanorequal` | `lessThanOrEqual` |

For `between` / `notBetween` supply both `formula1` (lower bound) and `formula2` (upper bound). For all other operators only `formula1` is needed.

## Formulas

`formula1` and `formula2` are free-text strings that become the `<formula1>` / `<formula2>` elements in the OOXML. They accept:

- **Inline list** — a comma-separated string wrapped in double quotes: `'"Alpha,Beta,Gamma"'`
- **Range reference** — an absolute reference to cells on any sheet: `'$E$2:$E$10'` or `'Lists!$A$1:$A$20'`
- **Literal value** — a number or date serial: `'100'`, `'42736'`
- **Formula expression** — any valid Excel formula: `'=INDIRECT("Lists!A:A")'`

::: tip List from a named range
For long or maintainable lists, define a named range (see [Named Ranges](./worksheets.md#named-ranges)) and reference it in `formula1`:
```abap
lo_validation->formula1 = 'Colours'.  " workbook-scoped named range
```
:::

## Cell Range

A single validation rule can cover a rectangular range of cells, not just one cell:

```abap
lo_validation->cell_column    = 'B'.
lo_validation->cell_row       = 2.
lo_validation->cell_column_to = 'B'.   " same column → entire column B rows 2-100
lo_validation->cell_row_to    = 100.
```

When `cell_column_to` / `cell_row_to` are blank the rule applies only to the single cell defined by `cell_column` / `cell_row`.

## Input Message

Display a tooltip when the user selects the cell:

```abap
lo_validation->showinputmessage = abap_true.
lo_validation->prompttitle      = 'Enter quantity'.
lo_validation->prompt           = 'Value must be between 1 and 999.'.
```

## Error Alert

Control what happens when invalid data is entered:

```abap
lo_validation->showerrormessage = abap_true.
lo_validation->errortitle       = 'Invalid input'.
lo_validation->error            = 'Please enter a number between 1 and 999.'.
lo_validation->errorstyle       = zcl_excel_data_validation=>c_style_stop.  " blocks entry
```

### Error Styles

| Constant | Value | Behaviour |
|---|---|---|
| `c_style_stop` | `stop` | Blocks invalid entry — user must correct or cancel |
| `c_style_warning` | `warning` | Warns but allows the user to accept anyway |
| `c_style_information` | `information` | Informational only — entry is always accepted |

## Allow Blank

`allowblank` (default `abap_false` after `constructor`) controls whether an empty cell passes validation. Set to `abap_true` to permit blanks:

```abap
lo_validation->allowblank = abap_true.
```

## Worked Examples

### Integer range 1–100

```abap
lo_validation->type        = zcl_excel_data_validation=>c_type_whole.
lo_validation->operator    = zcl_excel_data_validation=>c_operator_between.
lo_validation->formula1    = '1'.
lo_validation->formula2    = '100'.
lo_validation->allowblank  = abap_true.
lo_validation->errorstyle  = zcl_excel_data_validation=>c_style_stop.
lo_validation->errortitle  = 'Out of range'.
lo_validation->error       = 'Enter a whole number from 1 to 100.'.
lo_validation->cell_column = 'C'.
lo_validation->cell_row    = 3.
```

### Date not in the past

```abap
lo_validation->type      = zcl_excel_data_validation=>c_type_date.
lo_validation->operator  = zcl_excel_data_validation=>c_operator_greaterthanorequal.
lo_validation->formula1  = '=TODAY()'.
lo_validation->allowblank = abap_false.
lo_validation->errorstyle = zcl_excel_data_validation=>c_style_warning.
lo_validation->errortitle = 'Past date'.
lo_validation->error      = 'Delivery date should not be in the past.'.
lo_validation->cell_column    = 'D'.
lo_validation->cell_row       = 2.
lo_validation->cell_column_to = 'D'.
lo_validation->cell_row_to    = 500.
```

### Custom formula — unique values only

```abap
lo_validation->type     = zcl_excel_data_validation=>c_type_custom.
lo_validation->formula1 = '=COUNTIF($A$2:$A$500,A2)=1'.
lo_validation->showerrormessage = abap_true.
lo_validation->errortitle = 'Duplicate'.
lo_validation->error      = 'This value already exists in column A.'.
lo_validation->cell_column    = 'A'.
lo_validation->cell_row       = 2.
lo_validation->cell_column_to = 'A'.
lo_validation->cell_row_to    = 500.
```

## Adding to the Worksheet

Every `zcl_excel_worksheet` exposes a `data_validations` attribute of type `zcl_excel_data_validations`. Call `add()` to register each rule:

```abap
lo_worksheet->data_validations->add( lo_validation ).
```

The writer serialises all registered validations into the `<dataValidations>` element of the sheet XML automatically.

## Limitations

- Excel limits each worksheet to **65,534 data validation rules**.
- Cross-sheet list references (`Sheet2!$A$1:$A$10`) work in Excel desktop but may not evaluate in some viewers.
- The `c_type_custom` formula is written verbatim into the XML — ensure it uses column-relative references if the rule spans multiple rows.
- abap2xlsx does **not** validate the formula syntax itself; incorrect formulas silently produce no constraint in Excel.
