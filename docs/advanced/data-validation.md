# Data Validation

Advanced guide to implementing input validation rules in Excel worksheets with abap2xlsx.

## Understanding Data Validation

Data validation in Excel allows you to control what users can enter into cells by defining rules and constraints. The abap2xlsx library provides comprehensive support for creating data validation rules through the `zcl_excel_data_validation` class <cite>src/zcl_excel_data_validation.clas.xml:1-183</cite>.

## Basic Data Validation

### Creating Validation Rules

```abap
" Create basic data validation rule
METHOD create_basic_validation.
  DATA: lo_data_validation TYPE REF TO zcl_excel_data_validation.

  " Create validation for numeric range
  CREATE OBJECT lo_data_validation.
  lo_data_validation->type = zcl_excel_data_validation=>c_type_whole.
  lo_data_validation->operator = zcl_excel_data_validation=>c_operator_between.
  lo_data_validation->formula1 = '1'.
  lo_data_validation->formula2 = '100'.
  lo_data_validation->cell_column = 'B'.
  lo_data_validation->cell_row = 2.
  lo_data_validation->cell_column_to = 'B'.
  lo_data_validation->cell_row_to = 100.

  " Add to worksheet
  lo_worksheet->add_data_validation( lo_data_validation ).
ENDMETHOD.
```

### Validation Types

The data validation system supports multiple validation types <cite>src/zcl_excel_data_validation.clas.xml:100-139</cite>:

```abap
" Different validation types
METHOD create_validation_types.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Whole number validation
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_whole.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_greaterthan.
  lo_validation->formula1 = '0'.
  apply_validation_to_range( lo_validation 'C2:C100' ).

  " Decimal validation
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_decimal.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_between.
  lo_validation->formula1 = '0.00'.
  lo_validation->formula2 = '999.99'.
  apply_validation_to_range( lo_validation 'D2:D100' ).

  " Date validation
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_date.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_greaterthanorequal.
  lo_validation->formula1 = 'TODAY()'.
  apply_validation_to_range( lo_validation 'E2:E100' ).

  " Time validation
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_time.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_between.
  lo_validation->formula1 = '08:00'.
  lo_validation->formula2 = '18:00'.
  apply_validation_to_range( lo_validation 'F2:F100' ).
ENDMETHOD.
```

## Advanced Validation Rules

### List-Based Validation

```abap
" Create dropdown lists for data validation
METHOD create_list_validation.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Static list validation
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_list.
  lo_validation->formula1 = '"High,Medium,Low"'.
  lo_validation->showdropdown = abap_true.
  apply_validation_to_range( lo_validation 'G2:G100' ).

  " Dynamic list from range
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_list.
  lo_validation->formula1 = '$A$1:$A$10'.  " Reference to range
  lo_validation->showdropdown = abap_true.
  apply_validation_to_range( lo_validation 'H2:H100' ).

  " Named range list
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_list.
  lo_validation->formula1 = 'StatusList'.  " Named range
  lo_validation->showdropdown = abap_true.
  apply_validation_to_range( lo_validation 'I2:I100' ).
ENDMETHOD.
```

### Custom Formula Validation

```abap
" Create custom validation using formulas
METHOD create_custom_validation.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Custom formula validation
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_custom.
  lo_validation->formula1 = 'AND(LEN(J2)>=3,LEN(J2)<=10)'.  " Text length between 3-10
  apply_validation_to_range( lo_validation 'J2:J100' ).

  " Email format validation
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_custom.
  lo_validation->formula1 = 'AND(ISERROR(FIND(" ",K2)),LEN(K2)-LEN(SUBSTITUTE(K2,"@",""))=1)'.
  apply_validation_to_range( lo_validation 'K2:K100' ).

  " Unique value validation
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_custom.
  lo_validation->formula1 = 'COUNTIF($L$2:$L$100,L2)=1'.
  apply_validation_to_range( lo_validation 'L2:L100' ).
ENDMETHOD.
```

### Text Length Validation

```abap
" Validate text length constraints
METHOD create_text_length_validation.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Minimum text length
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_textlength.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_greaterthanorequal.
  lo_validation->formula1 = '5'.
  apply_validation_to_range( lo_validation 'M2:M100' ).

  " Maximum text length
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_textlength.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_lessthanorequal.
  lo_validation->formula1 = '50'.
  apply_validation_to_range( lo_validation 'N2:N100' ).

  " Text length range
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_textlength.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_between.
  lo_validation->formula1 = '3'.
  lo_validation->formula2 = '20'.
  apply_validation_to_range( lo_validation 'O2:O100' ).
ENDMETHOD.
```

## Validation Messages

### Input Messages

```abap
" Configure input messages for validation
METHOD configure_input_messages.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_whole.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_between.
  lo_validation->formula1 = '1'.
  lo_validation->formula2 = '100'.
  
  " Configure input message
  lo_validation->showinputmessage = abap_true.
  lo_validation->prompttitle = 'Enter Value'.
  lo_validation->prompt = 'Please enter a number between 1 and 100'.
  
  apply_validation_to_range( lo_validation 'P2:P100' ).
ENDMETHOD.
```

### Error Messages

The validation system supports different error styles <cite>src/zcl_excel_data_validation.clas.xml:86-99</cite>:

```abap
" Configure error messages and styles
METHOD configure_error_messages.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Stop error (prevents invalid input)
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_decimal.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_greaterthan.
  lo_validation->formula1 = '0'.
  lo_validation->showerrormessage = abap_true.
  lo_validation->errorstyle = zcl_excel_data_validation=>c_style_stop.
  lo_validation->errortitle = 'Invalid Input'.
  lo_validation->error = 'Value must be greater than 0'.
  apply_validation_to_range( lo_validation 'Q2:Q100' ).

  " Warning error (allows override)
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_whole.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_lessthan.
  lo_validation->formula1 = '1000'.
  lo_validation->showerrormessage = abap_true.
  lo_validation->errorstyle = zcl_excel_data_validation=>c_style_warning.
  lo_validation->errortitle = 'Large Value'.
  lo_validation->error = 'Are you sure you want to enter a value this large?'.
  apply_validation_to_range( lo_validation 'R2:R100' ).

  " Information error (informational only)
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_textlength.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_greaterthan.
  lo_validation->formula1 = '10'.
  lo_validation->showerrormessage = abap_true.
  lo_validation->errorstyle = zcl_excel_data_validation=>c_style_information.
  lo_validation->errortitle = 'Long Text'.
  lo_validation->error = 'This text is longer than recommended'.
  apply_validation_to_range( lo_validation 'S2:S100' ).
ENDMETHOD.
```

## Advanced Validation Scenarios

### Dependent Dropdowns

```abap
" Create cascading dropdown lists
METHOD create_dependent_dropdowns.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Primary dropdown (Categories)
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_list.
  lo_validation->formula1 = '"Electronics,Clothing,Books"'.
  lo_validation->showdropdown = abap_true.
  apply_validation_to_range( lo_validation 'T2:T100' ).

  " Secondary dropdown (depends on primary)
  CLEAR lo_validation.
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_list.
  lo_validation->formula1 = 'INDIRECT(T2)'.  " References named ranges
  lo_validation->showdropdown = abap_true.
  apply_validation_to_range( lo_validation 'U2:U100' ).
ENDMETHOD.
```

### Conditional Validation

```abap
" Apply validation based on other cell values
METHOD create_conditional_validation.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  " Validation depends on another cell
  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_custom.
  lo_validation->formula1 = 'IF(V2="Yes",AND(W2>=1,W2<=10),W2>=0)'.
  apply_validation_to_range( lo_validation 'W2:W100' ).
ENDMETHOD.
```

## Writer Integration

The data validation rules are processed by the Excel writer during file generation <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:1281-1393</cite>. The writer handles all validation attributes including type, operator, formulas, and message configurations.

## Utility Methods

```abap
" Helper method to apply validation to ranges
METHOD apply_validation_to_range.
  DATA: lv_range TYPE string,
        lt_range_parts TYPE TABLE OF string,
        lv_start_cell TYPE string,
        lv_end_cell TYPE string.

  " Parse range (e.g., 'A2:A100')
  SPLIT iv_range AT ':' INTO lv_start_cell lv_end_cell.
  
  " Extract column and row information
  extract_cell_coordinates( 
    iv_cell = lv_start_cell
    IMPORTING ev_column = io_validation->cell_column
              ev_row = io_validation->cell_row
  ).
  
  IF lv_end_cell IS NOT INITIAL.
    extract_cell_coordinates(
      iv_cell = lv_end_cell
      IMPORTING ev_column = io_validation->cell_column_to
                ev_row = io_validation->cell_row_to
    ).
  ENDIF.

  " Add validation to worksheet
  lo_worksheet->add_data_validation( io_validation ).
ENDMETHOD.
```

I'll continue from where the data validation guide left off. Here's the completion of that file:

## `docs/advanced/data-validation.md` (continued)

```markdown
## Complete Data Validation Example

### Form with Multiple Validation Rules

```abap
" Complete example: Employee data entry form
METHOD create_employee_form_validation.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zcl_excel_writer_2007.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Employee Form' ).

  " Create form headers
  create_form_headers( lo_worksheet ).

  " Employee ID validation (required, unique)
  create_employee_id_validation( lo_worksheet ).
  
  " Name validation (text length)
  create_name_validation( lo_worksheet ).
  
  " Email validation (format)
  create_email_validation( lo_worksheet ).
  
  " Department dropdown
  create_department_validation( lo_worksheet ).
  
  " Salary validation (numeric range)
  create_salary_validation( lo_worksheet ).
  
  " Start date validation (future dates only)
  create_date_validation( lo_worksheet ).

  " Generate Excel file
  CREATE OBJECT lo_writer.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
  
  MESSAGE 'Employee form with validation created successfully' TYPE 'S'.
ENDMETHOD.

METHOD create_salary_validation.
  DATA: lo_validation TYPE REF TO zcl_excel_data_validation.

  CREATE OBJECT lo_validation.
  lo_validation->type = zcl_excel_data_validation=>c_type_decimal.
  lo_validation->operator = zcl_excel_data_validation=>c_operator_between.
  lo_validation->formula1 = '30000'.
  lo_validation->formula2 = '200000'.
  lo_validation->showinputmessage = abap_true.
  lo_validation->prompttitle = 'Salary Range'.
  lo_validation->prompt = 'Enter annual salary between 30,000 and 200,000'.
  lo_validation->showerrormessage = abap_true.
  lo_validation->errorstyle = zcl_excel_data_validation=>c_style_stop.
  lo_validation->errortitle = 'Invalid Salary'.
  lo_validation->error = 'Salary must be between 30,000 and 200,000'.
  
  apply_validation_to_range( lo_validation 'F2:F1000' ).
ENDMETHOD.
```

## Integration with Worksheet System

The data validation system integrates seamlessly with the worksheet's data validation collection <cite>src/zcl_excel_worksheet.clas.abap:894-898</cite>. When you call `add_new_data_validation()`, it creates a new validation object and adds it to the worksheet's validation collection <cite>src/zcl_excel_data_validations.clas.abap:42-44</cite>.

The validation rules are then processed by the Excel writer during file generation <cite>src/zcl_excel_writer_2007.clas.locals_imp.abap:1281-1393</cite>, which handles all validation attributes including type, operator, formulas, and message configurations.

## Best Practices

### Design Guidelines

1. **User Experience**: Provide clear, helpful input and error messages
2. **Validation Logic**: Use appropriate validation types for data requirements
3. **Error Handling**: Choose appropriate error styles (stop, warning, information)
4. **Performance**: Limit validation rules to necessary ranges

### Implementation Guidelines

1. **Consistency**: Use consistent validation patterns across your workbook
2. **Testing**: Test validation rules with various input scenarios
3. **Documentation**: Document validation requirements for users
4. **Maintenance**: Keep validation rules updated with business requirements

## Next Steps

After mastering data validation:

- **[Password Protection](/advanced/password-protection)** - Secure your validated workbooks
- **[Macros](/advanced/macros)** - Automate validation processes
- **[Templates](/advanced/templates)** - Create reusable forms with validation

## Common Data Validation Patterns

### Quick Reference

```abap
" Create validation object
CREATE OBJECT lo_validation.

" Set validation type and operator
lo_validation->type = zcl_excel_data_validation=>c_type_whole.
lo_validation->operator = zcl_excel_data_validation=>c_operator_between.

" Configure formulas
lo_validation->formula1 = '1'.
lo_validation->formula2 = '100'.

" Add messages
lo_validation->showinputmessage = abap_true.
lo_validation->prompttitle = 'Enter Value'.
lo_validation->prompt = 'Please enter a number between 1 and 100'.

" Add to worksheet
lo_worksheet->add_data_validation( lo_validation ).
```

This guide covers the comprehensive data validation capabilities of abap2xlsx, enabling you to create robust, user-friendly Excel forms with sophisticated input validation and error handling.
