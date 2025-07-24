# Excel Formulas

Comprehensive guide to adding Excel formulas and calculations to your spreadsheets with abap2xlsx.

## Understanding Excel Formulas

### Basic Formula Concepts

Excel formulas in abap2xlsx are added using the `set_cell_formula` method. Formulas are stored separately from cell values and are evaluated by Excel when the file is opened.

```abap
" Basic formula example
lo_worksheet->set_cell_formula(
  ip_column = 'C'
  ip_row = 1
  ip_formula = 'A1+B1'
).

" Formula with cell value (for display before Excel calculates)
lo_worksheet->set_cell(
  ip_column = 'C'
  ip_row = 1
  ip_value = '0'  " Placeholder value
  ip_formula = 'A1+B1'
).
```

### Formula Syntax

Excel formulas in abap2xlsx follow standard Excel syntax:

```abap
" Arithmetic operations
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 1 ip_formula = 'A1+B1' ).      " Addition
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 2 ip_formula = 'A2-B2' ).      " Subtraction
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 3 ip_formula = 'A3*B3' ).      " Multiplication
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 4 ip_formula = 'A4/B4' ).      " Division
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 5 ip_formula = 'A5^2' ).       " Exponentiation

" Comparison operations
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 1 ip_formula = 'A1>B1' ).      " Greater than
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 2 ip_formula = 'A2=B2' ).      " Equal to
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 3 ip_formula = 'A3<>B3' ).     " Not equal to
```

## Common Excel Functions

### Mathematical Functions

```abap
" SUM function
lo_worksheet->set_cell_formula(
  ip_column = 'D'
  ip_row = 10
  ip_formula = 'SUM(A1:A9)'
).

" AVERAGE function
lo_worksheet->set_cell_formula(
  ip_column = 'E'
  ip_row = 10
  ip_formula = 'AVERAGE(B1:B9)'
).

" COUNT and COUNTA functions
lo_worksheet->set_cell_formula( ip_column = 'F' ip_row = 10 ip_formula = 'COUNT(C1:C9)' ).    " Count numbers
lo_worksheet->set_cell_formula( ip_column = 'G' ip_row = 10 ip_formula = 'COUNTA(C1:C9)' ).   " Count non-empty cells

" MIN and MAX functions
lo_worksheet->set_cell_formula( ip_column = 'H' ip_row = 10 ip_formula = 'MIN(A1:A9)' ).
lo_worksheet->set_cell_formula( ip_column = 'I' ip_row = 10 ip_formula = 'MAX(A1:A9)' ).

" ROUND function
lo_worksheet->set_cell_formula( ip_column = 'J' ip_row = 1 ip_formula = 'ROUND(A1/B1,2)' ).   " Round to 2 decimals
```

### Logical Functions

```abap
" IF function
lo_worksheet->set_cell_formula(
  ip_column = 'F'
  ip_row = 1
  ip_formula = 'IF(A1>100,"High","Low")'
).

" Nested IF functions
lo_worksheet->set_cell_formula(
  ip_column = 'G'
  ip_row = 1
  ip_formula = 'IF(A1>1000,"Very High",IF(A1>500,"High","Low"))'
).

" AND and OR functions
lo_worksheet->set_cell_formula( ip_column = 'H' ip_row = 1 ip_formula = 'IF(AND(A1>50,B1<100),"Valid","Invalid")' ).
lo_worksheet->set_cell_formula( ip_column = 'I' ip_row = 1 ip_formula = 'IF(OR(A1>1000,B1>1000),"Large","Small")' ).

" NOT function
lo_worksheet->set_cell_formula( ip_column = 'J' ip_row = 1 ip_formula = 'IF(NOT(A1=0),B1/A1,"N/A")' ).
```

### Text Functions

```abap
" CONCATENATE function (or & operator)
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 1 ip_formula = 'A1&" - "&B1' ).
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 1 ip_formula = 'CONCATENATE(A1," ",B1)' ).

" TEXT function for formatting
lo_worksheet->set_cell_formula( ip_column = 'F' ip_row = 1 ip_formula = 'TEXT(A1,"#,##0.00")' ).
lo_worksheet->set_cell_formula( ip_column = 'G' ip_row = 1 ip_formula = 'TEXT(TODAY(),"dd/mm/yyyy")' ).

" String manipulation functions
lo_worksheet->set_cell_formula( ip_column = 'H' ip_row = 1 ip_formula = 'LEFT(A1,5)' ).      " First 5 characters
lo_worksheet->set_cell_formula( ip_column = 'I' ip_row = 1 ip_formula = 'RIGHT(A1,3)' ).     " Last 3 characters
lo_worksheet->set_cell_formula( ip_column = 'J' ip_row = 1 ip_formula = 'MID(A1,3,4)' ).     " 4 chars starting at position 3
lo_worksheet->set_cell_formula( ip_column = 'K' ip_row = 1 ip_formula = 'LEN(A1)' ).         " Length of text
```

### Date and Time Functions

```abap
" Current date and time
lo_worksheet->set_cell_formula( ip_column = 'A' ip_row = 1 ip_formula = 'TODAY()' ).
lo_worksheet->set_cell_formula( ip_column = 'B' ip_row = 1 ip_formula = 'NOW()' ).

" Date calculations
lo_worksheet->set_cell_formula( ip_column = 'C' ip_row = 1 ip_formula = 'TODAY()+30' ).      " 30 days from today
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 1 ip_formula = 'DATEDIF(A1,B1,"D")' ). " Days between dates

" Date component extraction
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 1 ip_formula = 'YEAR(TODAY())' ).
lo_worksheet->set_cell_formula( ip_column = 'F' ip_row = 1 ip_formula = 'MONTH(TODAY())' ).
lo_worksheet->set_cell_formula( ip_column = 'G' ip_row = 1 ip_formula = 'DAY(TODAY())' ).

" WEEKDAY function
lo_worksheet->set_cell_formula( ip_column = 'H' ip_row = 1 ip_formula = 'WEEKDAY(TODAY())' ).
```

## Advanced Formula Techniques

### Array Formulas

```abap
" Array formula for multiple calculations
lo_worksheet->set_cell_formula(
  ip_column = 'D'
  ip_row = 1
  ip_formula = 'SUM(A1:A10*B1:B10)'  " Multiply arrays and sum
).

" SUMPRODUCT function
lo_worksheet->set_cell_formula(
  ip_column = 'E'
  ip_row = 1
  ip_formula = 'SUMPRODUCT(A1:A10,B1:B10)'
).
```

### Lookup Functions

```abap
" VLOOKUP function
lo_worksheet->set_cell_formula(
  ip_column = 'F'
  ip_row = 1
  ip_formula = 'VLOOKUP(A1,Table1,2,FALSE)'
).

" INDEX and MATCH combination
lo_worksheet->set_cell_formula(
  ip_column = 'G'
  ip_row = 1
  ip_formula = 'INDEX(C:C,MATCH(A1,B:B,0))'
).

" HLOOKUP function
lo_worksheet->set_cell_formula(
  ip_column = 'H'
  ip_row = 1
  ip_formula = 'HLOOKUP(A1,A1:Z5,3,FALSE)'
).
```

### Cross-Worksheet References

```abap
" Reference cells in other worksheets
lo_worksheet->set_cell_formula(
  ip_column = 'A'
  ip_row = 1
  ip_formula = 'Summary!B5'  " Reference cell B5 in Summary worksheet
).

" Reference ranges in other worksheets
lo_worksheet->set_cell_formula(
  ip_column = 'B'
  ip_row = 1
  ip_formula = 'SUM(Data!A1:A100)'
).

" Reference with spaces in sheet name
lo_worksheet->set_cell_formula(
  ip_column = 'C'
  ip_row = 1
  ip_formula = '''Sales Data''!C10'  " Use single quotes for sheet names with spaces
).
```

## Formula Management and Copying

### Copying Formulas with Relative References

The abap2xlsx library includes functionality for shifting formulas when copying cells:

```abap
" Example: Copy a formula from one cell to another with adjusted references
DATA: lv_original_formula TYPE string VALUE 'SUM(A1:A10)',
      lv_shifted_formula TYPE string.

" Shift formula 2 columns right and 3 rows down
lv_shifted_formula = zcl_excel_common=>shift_formula(
  iv_reference_formula = lv_original_formula
  iv_shift_cols = 2
  iv_shift_rows = 3
).
" Result: 'SUM(C4:C13)'

lo_worksheet->set_cell_formula(
  ip_column = 'F'
  ip_row = 5
  ip_formula = lv_shifted_formula
).
```

### Absolute vs Relative References

```abap
" Relative references (adjust when copied)
lo_worksheet->set_cell_formula( ip_column = 'C' ip_row = 1 ip_formula = 'A1*B1' ).

" Absolute references (don't adjust when copied)
lo_worksheet->set_cell_formula( ip_column = 'C' ip_row = 2 ip_formula = '$A$1*B2' ).

" Mixed references
lo_worksheet->set_cell_formula( ip_column = 'C' ip_row = 3 ip_formula = '$A1*B$1' ).  " Column A fixed, row 1 fixed
```

## Working with Named Ranges in Formulas

```abap
" Create named range first
DATA: lo_range TYPE REF TO zcl_excel_range.
lo_range = lo_excel->add_new_range( ).
lo_range->set_name( 'SalesData' ).
lo_range->set_value( 'Sheet1!$A$1:$A$100' ).

" Use named range in formulas
lo_worksheet->set_cell_formula(
  ip_column = 'B'
  ip_row = 1
  ip_formula = 'SUM(SalesData)'
).

lo_worksheet->set_cell_formula(
  ip_column = 'C'
  ip_row = 1
  ip_formula = 'AVERAGE(SalesData)'
).
```

## Error Handling in Formulas

### Formula Validation

```abap
" Add error checking to formulas
lo_worksheet->set_cell_formula(
  ip_column = 'D'
  ip_row = 1
  ip_formula = 'IF(ISERROR(A1/B1),"Division Error",A1/B1)'
).

" IFERROR function (Excel 2007+)
lo_worksheet->set_cell_formula(
  ip_column = 'E'
  ip_row = 1
  ip_formula = 'IFERROR(VLOOKUP(A1,Table1,2,FALSE),"Not Found")'
).

" Check for blank cells
lo_worksheet->set_cell_formula(
  ip_column = 'F'
  ip_row = 1
  ip_formula = 'IF(ISBLANK(A1),"No Data",A1*2)'
).
```

## Performance Considerations

### Efficient Formula Design

```abap
" Efficient: Use single formula for multiple calculations
lo_worksheet->set_cell_formula(
  ip_column = 'G'
  ip_row = 1
  ip_formula = 'SUMPRODUCT((A1:A100>50)*(B1:B100))'  " Sum B values where A > 50
).

" Less efficient: Multiple helper columns
" Better to combine logic into single formula when possible
```

### Formula Complexity

```abap
" Break complex formulas into manageable parts
" Instead of one very complex formula, use intermediate calculations

" Step 1: Calculate base value
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 1 ip_formula = 'A1*B1' ).

" Step 2: Apply discount
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 1 ip_formula = 'D1*(1-C1)' ).

" Step 3: Add tax
lo_worksheet->set_cell_formula( ip_column = 'F' ip_row = 1 ip_formula = 'E1*1.1' ).
```

## Complete Formula Example

### Sales Report with Calculations

```abap
" Complete example: Sales report with various formulas
METHOD create_sales_report_with_formulas.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Analysis' ).
  
  " Headers
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 1 ip_value = 'Product' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 1 ip_value = 'Quantity' ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 1 ip_value = 'Unit Price' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Discount %' ).
  lo_worksheet->set_cell( ip_column = 'E' ip_row = 1 ip_value = 'Subtotal' ).
  lo_worksheet->set_cell( ip_column = 'F' ip_row = 1 ip_value = 'Tax' ).
  lo_worksheet->set_cell( ip_column = 'G' ip_row = 1 ip_value = 'Total' ).

  " Sample data
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 2 ip_value = 'Laptop' ).
  lo_worksheet->set_cell( ip_column = 'B' ip_row = 2 ip_value = 5 ).
  lo_worksheet->set_cell( ip_column = 'C' ip_row = 2 ip_value = '999.99' ).
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 2 ip_value = '0.1' ).

  " Formulas for calculations
  " Subtotal = Quantity * Unit Price * (1 - Discount)
  lo_worksheet->set_cell_formula(
    ip_column = 'E'
    ip_row = 2
    ip_formula = 'B2*C2*(1-D2)'
  ).

  " Tax = Subtotal * 0.1 (10% tax)
  lo_worksheet->set_cell_formula(
    ip_column = 'F'
    ip_row = 2
    ip_formula = 'E2*0.1'
  ).

  " Total = Subtotal + Tax
  lo_worksheet->set_cell_formula(
    ip_column = 'G'
    ip_row = 2
    ip_formula = 'E2+F2'
  ).

  " Summary totals
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 10 ip_value = 'TOTALS:' ).
  lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 10 ip_formula = 'SUM(E2:E9)' ).
  lo_worksheet->set_cell_formula( ip_column = 'F' ip_row = 10 ip_formula = 'SUM(F2:F9)' ).
  lo_worksheet->set_cell_formula( ip_column = 'G' ip_row = 10 ip_formula = 'SUM(G2:G9)' ).

  " Average calculation
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 11 ip_value = 'AVERAGE:' ).
  lo_worksheet->set_cell_formula( ip_column = 'G' ip_row = 11 ip_formula = 'AVERAGE(G2:G9)' ).

  " Conditional summary
  lo_worksheet->set_cell( ip_column = 'A' ip_row = 12 ip_value = 'High Value Items (>$1000):' ).
  lo_worksheet->set_cell_formula( ip_column = 'G' ip_row = 12 ip_formula = 'COUNTIF(G2:G9,">1000")' ).

  DATA: lo_writer TYPE REF TO zif_excel_writer.
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
ENDMETHOD.
```

## Column Formulas for Tables

abap2xlsx supports column formulas for table structures, which allow you to define formulas that apply to entire columns within a table:

```abap
" Define column formula for table
DATA: ls_column_formula TYPE zcl_excel_worksheet=>mty_s_column_formula.

ls_column_formula-id = 1.
ls_column_formula-column = 3.  " Column C
ls_column_formula-formula = 'A{row}*B{row}'.  " {row} will be replaced with actual row number
ls_column_formula-table_top_left_row = 2.
ls_column_formula-table_bottom_right_row = 10.

" Apply column formula to cells
lo_worksheet->set_cell(
  ip_column = 'C'
  ip_row = 2
  ip_column_formula_id = 1
).
```

## Best Practices for Formulas

### Formula Design Principles

1. **Keep Formulas Simple**: Break complex calculations into multiple steps
2. **Use Named Ranges**: Make formulas more readable and maintainable
3. **Add Error Handling**: Always include error checking for division and lookups
4. **Document Complex Formulas**: Add comments explaining the business logic

### Performance Optimization

1. **Minimize Volatile Functions**: Avoid excessive use of NOW(), TODAY(), RAND()
2. **Use Efficient Lookup Methods**: INDEX/MATCH often performs better than VLOOKUP
3. **Limit Array Formulas**: Use them judiciously for large datasets
4. **Cache Intermediate Results**: Store calculated values rather than recalculating

### Formula Maintenance

1. **Consistent Reference Style**: Use consistent absolute/relative referencing
2. **Avoid Hard-Coded Values**: Use cell references or named constants
3. **Test Edge Cases**: Verify formulas work with empty cells, zeros, and errors
4. **Version Control**: Document formula changes and their business rationale

## Next Steps

After mastering Excel formulas:

- **[Charts and Graphs](/guide/charts)** - Visualize your calculated data
- **[Data Conversion](/guide/data-conversion)** - Efficiently populate worksheets with ABAP data
- **[ALV Integration](/guide/alv-integration)** - Convert ALV grids with formulas
- **[Performance Optimization](/guide/performance)** - Optimize formula-heavy workbooks

## Common Formula Patterns

### Quick Reference for Formula Operations

```abap
" Basic arithmetic
lo_worksheet->set_cell_formula( ip_column = 'C' ip_row = 1 ip_formula = 'A1+B1' ).

" Conditional logic
lo_worksheet->set_cell_formula( ip_column = 'D' ip_row = 1 ip_formula = 'IF(A1>100,"High","Low")' ).

" Aggregation functions
lo_worksheet->set_cell_formula( ip_column = 'E' ip_row = 1 ip_formula = 'SUM(A1:A10)' ).

" Cross-sheet references
lo_worksheet->set_cell_formula( ip_column = 'F' ip_row = 1 ip_formula = 'Summary!B5' ).
```

This guide covers the comprehensive formula capabilities of abap2xlsx. The formula system supports the full range of Excel functions and enables you to create dynamic, calculated spreadsheets that automatically update when data changes.
