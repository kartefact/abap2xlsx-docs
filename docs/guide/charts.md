# Charts and Graphs

Comprehensive guide to creating charts and visual data representations with abap2xlsx.

## Chart Architecture

Charts in abap2xlsx are created through the drawing system, where charts are treated as special drawing objects. The chart creation process involves the writer classes that generate the necessary XML structures for Excel charts <cite>src/zcl_excel_writer_2007.clas.abap:1425-1530</cite>.

```abap
" Basic chart creation
DATA: lo_excel TYPE REF TO zcl_excel,
      lo_worksheet TYPE REF TO zcl_excel_worksheet,
      lo_chart TYPE REF TO zcl_excel_drawing.

CREATE OBJECT lo_excel.
lo_worksheet = lo_excel->add_new_worksheet( ).

" Create chart drawing
lo_chart = lo_excel->add_new_drawing( ).
lo_chart->set_type( zcl_excel_drawing=>type_chart ).

" Position the chart
lo_chart->set_position(
  ip_from_row = 5
  ip_from_col = 'B'
  ip_to_row = 15
  ip_to_col = 'H'
).

" Add chart to worksheet
lo_worksheet->add_drawing( lo_chart ).
```

## Chart Types

### Bar Charts

```abap
" Create bar chart
DATA: lo_bar_chart TYPE REF TO zcl_excel_graph_bars.

CREATE OBJECT lo_bar_chart.

" Configure bar chart properties
lo_bar_chart->set_chart_type( zcl_excel_graph_bars=>c_chart_type_column ).
lo_bar_chart->set_data_range( 'A2:B10' ).
lo_bar_chart->set_title( 'Sales by Region' ).

" Set bar direction
lo_bar_chart->set_bar_direction( zcl_excel_graph_bars=>c_bar_dir_col ).

" Configure gap width between bars
lo_bar_chart->set_gap_width( '150' ).

" Add data series
lo_bar_chart->add_series(
  ip_categories_range = 'A2:A10'
  ip_values_range = 'B2:B10'
  ip_series_title = 'Sales Amount'
).
```

### Line Charts

```abap
" Create line chart
DATA: lo_line_chart TYPE REF TO zcl_excel_graph_line.

CREATE OBJECT lo_line_chart.

" Configure line chart properties
lo_line_chart->set_chart_type( zcl_excel_graph_line=>c_chart_type_line ).
lo_line_chart->set_data_range( 'A1:C12' ).
lo_line_chart->set_title( 'Monthly Trend Analysis' ).

" Configure line properties
lo_line_chart->set_smooth( abap_true ).  " Smooth lines
lo_line_chart->set_marker_symbol( zcl_excel_graph_line=>c_symbol_circle ).

" Add multiple data series
lo_line_chart->add_series(
  ip_categories_range = 'A2:A12'
  ip_values_range = 'B2:B12'
  ip_series_title = '2023 Sales'
).

lo_line_chart->add_series(
  ip_categories_range = 'A2:A12'
  ip_values_range = 'C2:C12'
  ip_series_title = '2024 Sales'
).
```

### Pie Charts

```abap
" Create pie chart
DATA: lo_pie_chart TYPE REF TO zcl_excel_graph_pie.

CREATE OBJECT lo_pie_chart.

" Configure pie chart properties
lo_pie_chart->set_chart_type( zcl_excel_graph_pie=>c_chart_type_pie ).
lo_pie_chart->set_data_range( 'A1:B6' ).
lo_pie_chart->set_title( 'Market Share Distribution' ).

" Set first slice angle
lo_pie_chart->set_first_slice_angle( 90 ).

" Configure data labels
lo_pie_chart->set_show_legend_key( abap_false ).
lo_pie_chart->set_show_value( abap_true ).
lo_pie_chart->set_show_category_name( abap_true ).
lo_pie_chart->set_show_percentage( abap_true ).

" Add data series
lo_pie_chart->add_series(
  ip_categories_range = 'A2:A6'
  ip_values_range = 'B2:B6'
  ip_series_title = 'Market Share'
).
```

## Chart Customization

### Chart Titles and Labels

```abap
" Set chart title
lo_chart->set_title( 'Quarterly Sales Performance' ).

" Configure axis titles
lo_chart->set_x_axis_title( 'Quarter' ).
lo_chart->set_y_axis_title( 'Sales Amount (€)' ).

" Set data label options
lo_chart->set_show_data_labels( abap_true ).
lo_chart->set_data_label_position( zcl_excel_chart=>c_label_pos_outside_end ).
```

### Chart Legends

```abap
" Configure legend
lo_chart->set_legend_position( zcl_excel_chart=>c_legend_pos_right ).
lo_chart->set_legend_overlay( abap_false ).

" Show/hide legend
lo_chart->set_show_legend( abap_true ).
```

### Chart Axes Configuration

The chart axis configuration is handled through the XML generation process <cite>src/zcl_excel_writer_2007.clas.abap:1781-1800</cite>:

```abap
" Configure chart axes
DATA: ls_axis TYPE zcl_excel_graph_bars=>ty_axis.

" Category axis (X-axis)
ls_axis-axid = '1001'.
ls_axis-type = zcl_excel_graph_bars=>c_catax.
ls_axis-orientation = 'minMax'.
ls_axis-delete = abap_false.
ls_axis-axpos = 'b'.  " Bottom position

lo_bar_chart->add_axis( ls_axis ).

" Value axis (Y-axis)
ls_axis-axid = '1002'.
ls_axis-type = zcl_excel_graph_bars=>c_valax.
ls_axis-orientation = 'minMax'.
ls_axis-axpos = 'l'.  " Left position
ls_axis-crosses = 'autoZero'.

lo_bar_chart->add_axis( ls_axis ).
```

## Advanced Chart Features

### Multiple Data Series

```abap
" Add multiple data series to a chart
METHOD add_multiple_series.
  " First series - Actual values
  lo_chart->add_series(
    ip_categories_range = 'A2:A13'
    ip_values_range = 'B2:B13'
    ip_series_title = 'Actual Sales'
  ).

  " Second series - Target values
  lo_chart->add_series(
    ip_categories_range = 'A2:A13'
    ip_values_range = 'C2:C13'
    ip_series_title = 'Target Sales'
  ).

  " Third series - Previous year
  lo_chart->add_series(
    ip_categories_range = 'A2:A13'
    ip_values_range = 'D2:D13'
    ip_series_title = 'Previous Year'
  ).
ENDMETHOD.
```

### Chart Formatting and Colors

```abap
" Customize chart appearance
lo_chart->set_chart_style( 2 ).  " Apply predefined style

" Configure plot area
lo_chart->set_plot_area_fill_color( 'F5F5F5' ).
lo_chart->set_plot_area_border_color( '808080' ).

" Set series colors
lo_chart->set_series_color( 
  ip_series_index = 1
  ip_color = '4472C4'  " Blue
).

lo_chart->set_series_color(
  ip_series_index = 2
  ip_color = 'E70000'  " Red
).
```

### Chart Data Labels

The chart data labels are configured through the XML structure <cite>src/zcl_excel_writer_2007.clas.abap:1747-1773</cite>:

```abap
" Configure data labels
lo_chart->set_show_legend_key( abap_false ).
lo_chart->set_show_value( abap_true ).
lo_chart->set_show_category_name( abap_false ).
lo_chart->set_show_series_name( abap_false ).
lo_chart->set_show_percentage( abap_false ).  " For pie charts
lo_chart->set_show_bubble_size( abap_false ).  " For bubble charts
```

## Chart Integration with Data

### Dynamic Chart Data

```abap
" Create chart with dynamic data ranges
METHOD create_dynamic_chart.
  DATA: lv_last_row TYPE i,
        lv_data_range TYPE string.

  " Determine data range dynamically
  lv_last_row = lo_worksheet->get_highest_row( ).
  lv_data_range = |A1:B{ lv_last_row }|.

  " Create chart with dynamic range
  lo_chart->set_data_range( lv_data_range ).
  
  " Set categories and values ranges
  DATA(lv_categories) = |A2:A{ lv_last_row }|.
  DATA(lv_values) = |B2:B{ lv_last_row }|.
  
  lo_chart->add_series(
    ip_categories_range = lv_categories
    ip_values_range = lv_values
    ip_series_title = 'Dynamic Data'
  ).
ENDMETHOD.
```

### Chart with Calculated Data

```abap
" Create chart using formula-based data
METHOD create_calculated_chart.
  " Add calculated columns for chart data
  lo_worksheet->set_cell( ip_column = 'D' ip_row = 1 ip_value = 'Growth %' ).
  
  " Add growth calculation formulas
  DATA: lv_row TYPE i VALUE 2.
  DO 12 TIMES.
    lo_worksheet->set_cell_formula(
      ip_column = 'D'
      ip_row = lv_row
      ip_formula = |IF(B{ lv_row - 1 }>0,(B{ lv_row }-B{ lv_row - 1 })/B{ lv_row - 1 }*100,0)|
    ).
    ADD 1 TO lv_row.
  ENDDO.

  " Create chart using calculated data
  lo_chart->add_series(
    ip_categories_range = 'A2:A13'
    ip_values_range = 'D2:D13'
    ip_series_title = 'Growth Rate'
  ).
ENDMETHOD.
```

## Chart Templates and Reusability

### Chart Template Class

```abap
" Create reusable chart templates
CLASS zcl_excel_chart_templates DEFINITION.
  PUBLIC SECTION.
    CLASS-METHODS: create_sales_chart
                     IMPORTING io_excel TYPE REF TO zcl_excel
                               io_worksheet TYPE REF TO zcl_excel_worksheet
                               iv_data_range TYPE string
                     RETURNING VALUE(ro_chart) TYPE REF TO zcl_excel_drawing,
                   create_trend_chart
                     IMPORTING io_excel TYPE REF TO zcl_excel
                               io_worksheet TYPE REF TO zcl_excel_worksheet
                               iv_data_range TYPE string
                     RETURNING VALUE(ro_chart) TYPE REF TO zcl_excel_drawing.
ENDCLASS.

CLASS zcl_excel_chart_templates IMPLEMENTATION.
  METHOD create_sales_chart.
    " Standard sales chart template
    ro_chart = io_excel->add_new_drawing( ).
    ro_chart->set_type( zcl_excel_drawing=>type_chart ).
    
    " Configure standard sales chart properties
    DATA(lo_bar_chart) = NEW zcl_excel_graph_bars( ).
    lo_bar_chart->set_chart_type( zcl_excel_graph_bars=>c_chart_type_column ).
    lo_bar_chart->set_data_range( iv_data_range ).
    lo_bar_chart->set_title( 'Sales Performance' ).
    
    " Apply standard formatting
    lo_bar_chart->set_gap_width( '150' ).
    ro_chart->set_chart_object( lo_bar_chart ).
  ENDMETHOD.

  METHOD create_trend_chart.
    " Standard trend chart template
    ro_chart = io_excel->add_new_drawing( ).
    ro_chart->set_type( zcl_excel_drawing=>type_chart ).
    
    " Configure standard trend chart properties
    DATA(lo_line_chart) = NEW zcl_excel_graph_line( ).
    lo_line_chart->set_chart_type( zcl_excel_graph_line=>c_chart_type_line ).
    lo_line_chart->set_data_range( iv_data_range ).
    lo_line_chart->set_title( 'Trend Analysis' ).
    
    " Apply standard formatting
    lo_line_chart->set_smooth( abap_true ).
    ro_chart->set_chart_object( lo_line_chart ).
  ENDMETHOD.
ENDCLASS.
```

## Complete Chart Example

### Dashboard with Multiple Charts

```abap
" Complete example: Create dashboard with multiple chart types
METHOD create_chart_dashboard.
  DATA: lo_excel TYPE REF TO zcl_excel,
        lo_worksheet TYPE REF TO zcl_excel_worksheet,
        lo_writer TYPE REF TO zif_excel_writer.

  CREATE OBJECT lo_excel.
  lo_worksheet = lo_excel->add_new_worksheet( ).
  lo_worksheet->set_title( 'Sales Dashboard' ).

  " Add sample data
  populate_sample_data( lo_worksheet ).

  " Create bar chart for regional sales
  DATA(lo_bar_chart) = create_regional_sales_chart( lo_excel ).
  lo_bar_chart->set_position(
    ip_from_row = 2
    ip_from_col = 'E'
    ip_to_row = 12
    ip_to_col = 'K'
  ).
  lo_worksheet->add_drawing( lo_bar_chart ).

  " Create line chart for monthly trends
  DATA(lo_line_chart) = create_monthly_trend_chart( lo_excel ).
  lo_line_chart->set_position(
    ip_from_row = 14
    ip_from_col = 'E'
    ip_to_row = 24
    ip_to_col = 'K'
  ).
  lo_worksheet->add_drawing( lo_line_chart ).

  " Create pie chart for product mix
  DATA(lo_pie_chart) = create_product_mix_chart( lo_excel ).
  lo_pie_chart->set_position(
    ip_from_row = 2
    ip_from_col = 'L'
    ip_to_row = 12
    ip_to_col = 'R'
  ).
  lo_worksheet->add_drawing( lo_pie_chart ).

  " Generate Excel file
  CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
  DATA(lv_file) = lo_writer->write_file( lo_excel ).
ENDMETHOD.
```

## Performance Considerations

### Chart Optimization

```abap
" Optimize chart performance
METHOD optimize_chart_performance.
  " 1. Limit data points - too many points can slow rendering
  DATA: lv_max_points TYPE i VALUE 100.
  
  " 2. Use appropriate chart types for data size
  IF lines( lt_data ) > 1000.
    " Use line charts instead of scatter plots for large datasets
    lo_chart->set_chart_type( zcl_excel_graph_line=>c_chart_type_line ).
  ENDIF.
  
  " 3. Minimize chart complexity
  " Avoid too many data series in one chart
  IF lines( lt_series ) > 5.
    " Consider splitting into multiple charts
  ENDIF.
  
  " 4. Cache chart objects when creating multiple similar charts
  IF mo_cached_chart_template IS NOT BOUND.
    mo_cached_chart_template = create_standard_chart_template( ).
  ENDIF.
ENDMETHOD.
```

### Memory Management

```abap
" Proper cleanup for chart objects
METHOD cleanup_chart_objects.
  " Clear chart references
  CLEAR: lo_chart, lo_bar_chart, lo_line_chart, lo_pie_chart.
  
  " Clear drawing collections
  IF lo_drawings IS BOUND.
    lo_drawings->clear( ).
  ENDIF.
ENDMETHOD.
```

## Troubleshooting Charts

### Common Chart Issues

```abap
" Debug chart creation issues
METHOD debug_chart_issues.
  " 1. Verify data range exists
  DATA: lv_max_row TYPE i,
        lv_max_col TYPE i.
  
  lv_max_row = lo_worksheet->get_highest_row( ).
  lv_max_col = lo_worksheet->get_highest_column( ).
  
  IF lv_max_row < 2 OR lv_max_col < 2.
    MESSAGE 'Insufficient data for chart creation' TYPE 'W'.
    RETURN.
  ENDIF.
  
  " 2. Check for empty data ranges
  DATA(lv_test_value) = lo_worksheet->get_cell( ip_column = 'A' ip_row = 2 ).
  IF lv_test_value IS INITIAL.
    MESSAGE 'Chart data range appears to be empty' TYPE 'W'.
  ENDIF.
  
  " 3. Validate chart positioning
  IF iv_from_row >= iv_to_row OR iv_from_col >= iv_to_col.
    MESSAGE 'Invalid chart position coordinates' TYPE 'E'.
  ENDIF.
ENDMETHOD.
```

## Next Steps

After mastering charts and graphs:

- **[Images and Drawings](/guide/images)** - Add images and other visual elements
- **[Data Conversion](/guide/data-conversion)** - Efficiently populate charts with ABAP data
- **[ALV Integration](/guide/alv-integration)** - Create charts from ALV data
- **[Performance Optimization](/guide/performance)** - Optimize workbooks with multiple charts

## Common Chart Patterns

### Quick Reference for Chart Operations

```abap
" Create basic chart
DATA(lo_chart) = lo_excel->add_new_drawing( ).
lo_chart->set_type( zcl_excel_drawing=>type_chart ).

" Position chart
lo_chart->set_position(
  ip_from_row = 5
  ip_from_col = 'B'
  ip_to_row = 15
  ip_to_col = 'H'
).

" Add to worksheet
lo_worksheet->add_drawing( lo_chart ).
```

This guide covers the comprehensive charting capabilities of abap2xlsx. The chart system integrates with Excel's native charting engine to provide professional data visualizations that enhance your reports and dashboards.
