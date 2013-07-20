---
layout: default
title: Chart Method
---
### <a name="chart" class="anchor" href="#chart"><span class="octicon octicon-link" /></a>SYNOPSIS

To create a simple Excel file with a chart using WriteXLSX:

    require 'write_xlsx'

    workbook  = WriteXLSX.new('chart.xlsx')
    worksheet = workbook.add_worksheet

    # Add the worksheet data the chart refers to.
    data = [
        [ 'Category', 2, 3, 4, 5, 6, 7 ],
        [ 'Value',    1, 4, 5, 2, 1, 5 ]

    ]

    worksheet.write('A1', data)

    # Add a worksheet chart.
    chart = workbook.add_chart(:type => 'column')

    # Configure the chart.
    chart.add_series(
        categories => '=Sheet1!$A$2:$A$7',
        values     => '=Sheet1!$B$2:$B$7'
    )

### <a name="description" class="anchor" href="#description"><span class="octicon octicon-link" /></a>DESCRIPTION

The Chart module is an abstract base class for modules that implement charts
in WriteXLSX. The information below is applicable to all of the available subclasses.

The Chart module isn't used directly.
A chart object is created via the Workbook
[`add_chart()`](workbook.html#add_chart) method where the chart type is specified:

    chart = workbook.add_chart(:type => 'column')

Currently the supported chart types are:

* area
    <p>Creates an Area (filled line) style chart. See WriteXLSX::Chart::Area.</p>

* bar
    <p>Creates a Bar style (transposed histogram) chart. See WriteXLSX::Chart::Bar.</p>

* column
    <p>Creates a column style (histogram) chart. See WriteXLSX::Chart::Column.</p>

* line
    <p>Creates a Line style chart. See WriteXLSX::Chart::Line.</p>

* pie
    <p>Creates a Pie style chart. See WriteXLSX::Chart::Pie.</p>

* scatter
    <p>Creates a Scatter style chart. See WriteXLSX::Chart::Scatter.</p>

* stock
    <p>Creates a Stock style chart. See WriteXLSX::Chart::Stock.</p>

* radar
    <p>Creates a Radar style chart. See WriteXLSX::Chart::Radar.</p>

Chart subtypes are also supported in some cases:

    workbook.add_chart(:type => 'bar', :subtype => 'stacked')

The currently available subtypes are:

    area
        stacked
        percent_stacked

    bar
        stacked
        percent_stacked

    column
        stacked
        percent_stacked

    scatter
        straight_with_markers
        straight
        smooth_with_markers
        smooth

    radar
        with_markers
        filled

More charts and sub-types will be supported in time.

### <a name="chart_methods" class="anchor" href="#chart_methods"><span class="octicon octicon-link" /></a>CHART METHODS

Methods that are common to all chart types are documented below.
See the documentation for each of the above chart modules for chart specific information.

* [add_series](#add_series)
* [set_x_axis](#set_x_axis)
* [set_y_axis](#set_y_axis)
* [set_x2_axis](#set_x2_axis)
* [set_y2_axis](#set_y2_axis)
* [set_size](#set_size)
* [set_title](#set_title)
* [set_legend](#set_legend)
* [set_chartarea](#set_chartarea)
* [set_plotarea](#set_plotarea)
* [set_style](#set_style)
* [set_table](#set_table)
* [set_up_down_bars](#set_up_down_bars)
* [set_drop_lines](#set_drop_lines)
* [set_high_low_lines](#set_high_low_lines)
* [show_blanks_as](#show_blanks_as)
* [show_hidden_data](#show_hidden_data)

#### <a name="add_series" class="anchor" href="#add_series"><span class="octicon octicon-link" /></a>add_series()

In an Excel chart a "series" is a collection of information such as values,
X axis labels and the formatting that define which data is plotted.

With an WriteXLSX chart object the `add_series()` method is used
to set the properties for a series:

    chart.add_series(
        :categories => '=Sheet1!$A$2:$A$10', # Optional.
        :values     => '=Sheet1!$B$2:$B$10', # Required.
        :line       => { :color => 'blue' }
    )

The properties that can be set are:

##### <a name="values" class="anchor" href="#values"><span class="octicon octicon-link" /></a>:values
This is the most important property of a series and must be set for every
chart object.
It links the chart with the worksheet data that it displays.
A formula or array can be used for the data range, see below.

##### <a name="categories" class="anchor" href="#categories"><span class="octicon octicon-link" /></a>:categories
This sets the chart category labels.
The category is more or less the same as the X axis.
In most chart types the `:categories` property is optional and the chart
will just assume a sequential series from 1 .. n.

##### <a name="name" class="anchor" href="#name"><span class="octicon octicon-link" /></a>:name
Set the name for the series.
The name is displayed in the chart legend and in the formula bar.
The name property is optional and if it isn't supplied it will default to Series 1 .. n.

##### <a name="line" class="anchor" href="#line"><span class="octicon octicon-link" /></a>:line
Set the properties of the series line type such as colour and width.
See the [CHART FORMATTING][] section below.

##### <a name="border" class="anchor" href="#border"><span class="octicon octicon-link" /></a>:border
Set the border properties of the series such as colour and style.
See the [CHART FORMATTING][] section below.

##### <a name="fill" class="anchor" href="#fill"><span class="octicon octicon-link" /></a>:fill
Set the fill properties of the series such as colour.
See the [CHART FORMATTING][] section below.

##### <a name="marker" class="anchor" href="#marker"><span class="octicon octicon-link" /></a>:marker
Set the properties of the series marker such as style and colour.
See the [SERIES OPTIONS][] section below.

##### <a name="trendline" class="anchor" href="#trendline"><span class="octicon octicon-link" /></a>:trendline
Set the properties of the series trendline such as linear,
polynomial and moving average types.
See the [SERIES OPTIONS][] section below.

##### <a name="smooth" class="anchor" href="#smooth"><span class="octicon octicon-link" /></a>:smooth
The smooth option is used to set the smooth property of a line series.
See the [SERIES OPTIONS][] section below.

##### <a name="y_error_bars" class="anchor" href="#y_error_bars"><span class="octicon octicon-link" /></a>:y_error_bars
Set vertical error bounds for a chart series.
See the [SERIES OPTIONS][] section below.

##### <a name="x_error_bars" class="anchor" href="#x_error_bars"><span class="octicon octicon-link" /></a>:x_error_bars
Set horizontal error bounds for a chart series.
See the [SERIES OPTIONS][] section below.

##### <a name="data_labels" class="anchor" href="#data_labels"><span class="octicon octicon-link" /></a>:data_labels
Set data labels for the series.
See the [SERIES OPTIONS][] section below.

#### <a name="points" class="anchor" href="#points"><span class="octicon octicon-link" /></a>:points
Set properties for individual points in a series.
See the [SERIES OPTIONS][] section below.

##### <a name="invert_if_negative" class="anchor" href="#invert_if_negative"><span class="octicon octicon-link" /></a>:invert_if_negative
Invert the fill colour for negative values.
Usually only applicable to column and bar charts.

##### <a name="overlap" class="anchor" href="#overlap"><span class="octicon octicon-link" /></a>:overlap
Set the overlap between series in a Bar/Column chart.
The range is +/- 100. Default is 0.

    overlap => 20,

Note, it is only necessary to apply this property to one series of the chart.

##### <a name="gap" class="anchor" href="#gap"><span class="octicon octicon-link" /></a>:gap
Set the gap between series in a Bar/Column chart.
The range is 0 to 500. Default is 150.

    gap => 200,

Note, it is only necessary to apply this property to one series of the chart.

The categories and values can take either a range formula such as
`=Sheet1!$A$2:$A$7` or, more usefully when generating the range
programmatically, an array with zero indexed row/column values:

     [ sheetname, row_start, row_end, col_start, col_end ]

The following are equivalent:

    chart.add_series(:categories => '=Sheet1!$A$2:$A$7'     ) # Same as ...
    chart.add_series(:categories => [ 'Sheet1', 1, 6, 0, 0 ]) # Zero-indexed.

You can add more than one series to a chart.
In fact, some chart types such as stock require it.
The series numbering and order in the Excel chart will be the same as the order
in which they are added in WriteXLSX.

    # Add the first series.
    chart.add_series(
        :categories => '=Sheet1!$A$2:$A$7',
        :values     => '=Sheet1!$B$2:$B$7',
        :name       => 'Test data series 1'
    )

    # Add another series. Same categories. Different range values.
    chart.add_series(
        :categories => '=Sheet1!$A$2:$A$7',
        :values     => '=Sheet1!$C$2:$C$7',
        :name       => 'Test data series 2'
    )

#### <a name="set_x_axis" class="anchor" href="#set_x_axis"><span class="octicon octicon-link" /></a>set_x_axis()

The `set_x_axis()` method is used to set properties of the X axis.

    chart.set_x_axis(:name => 'Quarterly results')

The properties that can be set are:

    :name
    :name_font
    :num_font
    :num_format
    :min
    :max
    :minor_unit
    :major_unit
    :crossing
    :reverse
    :log_base
    :label_position
    :major_gridlines
    :minor_gridlines
    :visible

These are explained below.
Some properties are only applicable to value or category axes, as indicated.
See
"Value and Category Axes"
for an explanation of Excel's distinction between the axis types.

##### <a name="set_x_axis_name" class="anchor" href="#set_x_axis_name"><span class="octicon octicon-link" /></a>:name
Set the name (title or caption) for the axis. The name is displayed below the X axis.
The name property is optional. The default is to have no axis name.
(Applicable to category and value axes).

    chart.set_x_axis(:name => 'Quarterly results')

The name can also be a formula such as `=Sheet1!$A$1`.

##### <a name="set_x_axis_name_font" class="anchor" href="#set_x_font_name_font"><span class="octicon octicon-link" /></a>:name_font
Set the font properties for the axis title.
(Applicable to category and value axes).

    chart.set_x_axis(:name_font => {:name => 'Arial', :size => 10})

See the [CHART FONTS][] section below.

##### <a name="set_x_axis_num_font" class="anchor" href="#set_x_axis_num_font"><span class="octicon octicon-link" /></a>:num_font
Set the font properties for the axis numbers.
(Applicable to category and value axes).

    chart.set_x_axis(:num_font => {:bold => 1, :italic => 1})

See the [CHART FONTS][] section below.

##### <a name="set_x_axis_num_format" class="anchor" href="#set_x_axis_num_format"><span class="octicon octicon-link" /></a>:num_format
Set the number format for the axis.
(Applicable to category and value axes).

    chart.set_x_axis(:num_format => '#,##0.00')
    chart.set_y_axis(:num_format => '0.00%'   )

The number format is similar to the Worksheet Cell Format `:num_format` apart
from the fact that a format index cannot be used.
The explicit format string must be used as show above.
See
"[set_num_format()"](format.html#set_num_format)
in WriteXLSX for more information.

##### <a name="set_x_axis_min" class="anchor" href="#set_x_axis_min"><span class="octicon octicon-link" /></a>:min
Set the minimum value for the axis range.
(Applicable to value axes only.)

    chart.set_x_axis(:min => 20)

##### <a name="set_x_axis_max" class="anchor" href="#set_x_axis_max"><span class="octicon octicon-link" /></a>:max
Set the maximum value for the axis range.
(Applicable to value axes only.)

    chart.set_x_axis(:max => 80)

##### <a name="set_x_axis_minor_unit" class="anchor" href="#set_x_axis_minor_unit"><span class="octicon octicon-link" /></a>:minor_unit
Set the increment of the minor units in the axis range.
(Applicable to value axes only.)

    chart.set_x_axis(:minor_unit => 0.4)

##### <a name="set_x_axis_major_unit" class="anchor" href="#set_x_axis_major_unit"><span class="octicon octicon-link" /></a>:major_unit
Set the increment of the major units in the axis range.
(Applicable to value axes only.)

    chart.set_x_axis(:major_unit => 2)

##### <a name="set_x_axis_crossing" class="anchor" href="#set_x_axis_crossing"><span class="octicon octicon-link" /></a>:crossing
Set the position where the y axis will cross the x axis.
(Applicable to category and value axes.)

The crossing value can either be the string 'max' to set the crossing
at the maximum axis value or a numeric value.

    chart.set_x_axis(:crossing => 3)
    # or
    chart.set_x_axis(:crossing => 'max')

For category axes the numeric value must be an integer to represent
the category number that the axis crosses at.
For value axes it can have any value associated with the axis.

If crossing is omitted (the default) the crossing will be set automatically
by Excel based on the chart data.

##### <a name="set_x_axis_reverse" class="anchor" href="#set_x_axis_reverse"><span class="octicon octicon-link" /></a>:reverse
Reverse the order of the axis categories or values.
(Applicable to category and value axes.)

    chart.set_x_axis(:reverse => 1)

##### <a name="set_x_axis_log_base" class="anchor" href="#set_x_axis_log_base"><span class="octicon octicon-link" /></a>:log_base
Set the log base of the axis range.
(Applicable to value axes only.)

    chart.set_x_axis(:log_base => 10)

##### <a name="set_x_axis_label_position" class="anchor" href="#set_x_axis_label_position"><span class="octicon octicon-link" /></a>:label_position
Set the "Axis labels" position for the axis.
The following positions are available:

    next_to (the default)
    high
    low
    none

##### <a name="set_x_axis_major_gridlines" class="anchor" href="#set_x_axis_major_gridlines"><span class="octicon octicon-link" /></a>:major_gridlines
Configure the major gridlines for the axis. The available properties are:

    :visible
    :line

For example:

    chart.set_x_axis(
        :major_gridlines => {
            :visible => 1,
            :line    => { :color => 'red', :width => 1.25, :dash_type => 'dash' }
        }
    )

The visible property is usually on for the X-axis
but it depends on the type of chart.

The line property sets the gridline properties such as colour and width.
See the [CHART FORMATTING][] section below.

##### <a name="set_x_axis_minor_gridlines" class="anchor" href="#set_x_axis_minor_gridlines"><span class="octicon octicon-link" /></a>:minor_gridlines
This takes the same options as major_gridlines above.

The minor gridline visible property is off by default for all chart types.

##### <a name="set_x_axis_visible" class="anchor" href="#set_x_axis_visible"><span class="octicon octicon-link" /></a>:visible
Configure the visibility of the axis.

    chart.set_x_axis(:visible => 0)

More than one property can be set in a call to `set_x_axis()`:

    chart.set_x_axis(
        :name => 'Quarterly results',
        :min  => 10,
        :max  => 80
    )

#### <a name="set_y_axis" class="anchor" href="#set_y_axis"><span class="octicon octicon-link" /></a>set_y_axis()

The `set_y_axis()` method is used to set properties of the Y axis.
The properties that can be set are the same as for `set_x_axis`, see above.

#### <a name="set_x2_axis" class="anchor" href="#set_x2_axis"><span class="octicon octicon-link" /></a>set_x2_axis()

The `set_x2_axis()` method is used to set properties of the secondary X axis.
The properties that can be set are the same as for [set_x_axis()][], see above.
The default properties for this axis are:

    :label_position => 'none',
    :crossing       => 'max',
    :visible        => 0,

#### <a name="set_y2_axis" class="anchor" href="#set_y2_axis"><span class="octicon octicon-link" /></a>set_y2_axis()

The `set_y2_axis()` method is used to set properties of the secondary Y axis.
The properties that can be set are the same as for [set_x_axis()][], see above.
The default properties for this axis are:

    :major_gridlines => { :visible => 0 }

#### <a name="set_size" class="anchor" href="#set_size"><span class="octicon octicon-link" /></a>set_size()

The `set_size()` method is used to set the dimensions of the chart.
The size properties that can be set are:

     :width
     :height
     :x_scale
     :y_scale
     :x_offset
     :y_offset

The width and height are in pixels.
The default chart width is 480 pixels and the default height is 288 pixels.
The size of the chart can be modified by setting the width and height
or by setting the `:x_scale` and `:y_scale`:

    chart.set_size(:width => 720, :height => 576)

    # Same as:

    chart.set_size(:x_scale => 1.5, :y_scale => 2)

The `:x_offset` and `:y_offset` position the top left corner of the chart
in the cell that it is inserted into.

Note: the `:x_scale`, `:y_scale`, `:x_offset` and `:y_offset` parameters
can also be set via the [insert_chart()][] method:

    worksheet.insert_chart('E2', chart, 2, 4, 1.5, 2)

#### <a name="set_title" class="anchor" href="#set_title"><span class="octicon octicon-link" /></a>set_title()

The `set_title()` method is used to set properties of the chart title.

    chart.set_title(:name => 'Year End Results')

The properties that can be set are:

##### <a name="set_title_name" class="anchor" href="#set_title_name"><span class="octicon octicon-link" /></a>:name
Set the name (title) for the chart.
The name is displayed above the chart.
The name can also be a formula such as `=Sheet1!$A$1`.
The name property is optional.
The default is to have no chart title.

##### <a name="set_title_name_font" class="anchor" href="#set_title_name_font"><span class="octicon octicon-link" /></a>:name_font
Set the font properties for the chart title.
See the [CHART FONTS][] section below.

#### <a name="set_legend" class="anchor" href="#set_legend"><span class="octicon octicon-link" /></a>:set_legend()

The `set_legend()` method is used to set properties of the chart legend.

    chart.set_legend(:position => 'none')

The properties that can be set are:

##### <a name="set_legend_position" class="anchor" href="#set_legend_position"><span class="octicon octicon-link" /></a>:position
Set the position of the chart legend.

    chart.set_legend(:position => 'bottom')

The default legend position is right.
The available positions are:

    none
    top
    bottom
    left
    right
    overlay_left
    overlay_right

##### <a name="set_legend_delete_series" class="anchor" href="#set_legend_delete_series"><span class="octicon octicon-link" /></a>:delete_series
This allows you to remove 1 or more series from the the legend
(the series will still display on the chart).
This property takes an array ref as an argument and the series are zero indexed:

    # Delete/hide series index 0 and 2 from the legend.
    chart.set_legend(:delete_series => [0, 2])

#### <a name="set_chartarea" class="anchor" href="#set_chartarea"><span class="octicon octicon-link" /></a>set_chartarea()

The `set_chartarea()` method is used to set the properties of the chart area.

    chart.set_chartarea(
        :border => { :none  => 1 },
        :fill   => { :color => 'red' }
    )

The properties that can be set are:

##### <a name="set_chartarea_border" class="anchor" href="#set_chartarea_border"><span class="octicon octicon-link" /></a>:border
Set the border properties of the chartarea such as colour and style.
See the [CHART FORMATTING][] section below.

##### <a name="set_chartarea_fill" class="anchor" href="#set_chartarea_fill"><span class="octicon octicon-link" /></a>:fill
Set the fill properties of the chartarea such as colour.
See the [CHART FORMATTING][] section below.

#### <a name="set_plotarea" class="anchor" href="#set_plotarea"><span class="octicon octicon-link" /></a>set_plotarea()

The `set_plotarea()` method is used to set properties of the plot area of a chart.

    chart.set_plotarea(
        :border => { :color => 'yellow', :width => 1, :dash_type => 'dash' },
        :fill   => { :color => '#92D050' }
    )

The properties that can be set are:

##### <a name="set_plotarea_border" class="anchor" href="#set_plotarea_border"><span class="octicon octicon-link" /></a>:border
Set the border properties of the plotarea such as colour and style.
See the [CHART FORMATTING][] section below.

##### <a name="set_plotarea_fill" class="anchor" href="#set_plotarea_fill"><span class="octicon octicon-link" /></a>:fill
Set the fill properties of the plotarea such as colour.
See the [CHART FORMATTING][] section below.

#### <a name="set_style" class="anchor" href="#set_style"><span class="octicon octicon-link" /></a>set_style()

The `set_style()` method is used to set the style of the chart to one of
the 42 built-in styles available on the 'Design' tab in Excel:

    chart.set_style(4)

The default style is 2.

#### <a name="set_table" class="anchor" href="#set_table"><span class="octicon octicon-link" /></a>set_table()

The `set_table()` method adds a data table below the horizontal axis with
the data used to plot the chart.

    chart.set_table

The available options, with default values are:

    :vertical   => 1,    # Display vertical lines in the table.
    :horizontal => 1,    # Display horizontal lines in the table.
    :outline    => 1,    # Display an outline in the table.
    :show_keys  => 0     # Show the legend keys with the table data.

The data table can only be shown with Bar, Column, Line, Area and stock charts.

#### <a name="set_up_down_bars" class="anchor" href="#set_up_down_bars"><span class="octicon octicon-link" /></a>set_up_down_bars()

The `set_up_down_bars()` method adds Up-Down bars to Line charts to indicate
the difference between the first and last data series.

    chart.set_up_down_bars

It is possible to format the up and down bars to add fill and border properties
if required.
See the [CHART FORMATTING][] section below.

    chart.set_up_down_bars(
        :up   => { :fill => { :color => 'green' } },
        :down => { :fill => { :color => 'red' } }
    )

Up-down bars can only be applied to Line charts and to Stock charts (by default).

#### <a name="set_drop_lines" class="anchor" href="#set_drop_lines"><span class="octicon octicon-link" /></a>set_drop_lines

The `set_drop_lines()` method adds Drop Lines to charts to show the Category
value of points in the data.

    chart.set_drop_lines

It is possible to format the Drop Line line properties if required.
See the [CHART FORMATTING][] section below.

    chart.set_drop_lines(:line => { :color => 'red', :dash_type => 'square_dot' } )

Drop Lines are only available in Line, Area and Stock charts.

#### <a name="set_high_low_lines" class="anchor" href="#set_high_low_lines"><span class="octicon octicon-link" /></a>set_high_low_lines

The `set_high_low_lines()` method adds High-Low lines to charts to show the
maximum and minimum values of points in a Category.

    chart.set_high_low_lines

It is possible to format the High-Low Line line properties if required.
See the [CHART FORMATTING][] section below.

    chart.set_high_low_lines( :line => { :color => 'red' } )

High-Low Lines are only available in Line and Stock charts.

#### <a name="show_blanks_as" class="anchor" href="#show_blanks_as"><span class="octicon octicon-link" /></a>show_blanks_as()

The `show_blanks_as()` method controls how blank data is displayed in a chart.

    chart.show_blanks_as( 'span' )

The available options are:

        gap    # Blank data is shown as a gap. The default.
        zero   # Blank data is displayed as zero.
        span   # Blank data is connected with a line.

#### <a name="sjhow_hidden_data" class="anchor" href="#show_hidden_data"><span class="octicon octicon-link" /></a>show_hidden_data()

Display data in hidden rows or columns on the chart.

    chart.show_hidden_data

### <a name="series_options" class="anchor" href="#series_options"><span class="octicon octicon-link" /></a>SERIES OPTIONS

This section details the following properties of `add_series()` in more detail:

    :marker
    :trendline
    :y_error_bars
    :x_error_bars
    :data_labels
    :points
    :smooth

##### <a name="series_marker" class="anchor" href="#series_marker"><span class="octicon octicon-link" /></a>:marker

The marker format specifies the properties of the markers used to distinguish
series on a chart. In general only Line and Scatter chart types and trendlines
use markers.

The following properties can be set for marker formats in a chart.

    :type
    :size
    :border
    :fill

The type property sets the type of marker that is used with a series.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :marker     => { :type => 'diamond' }
    )

The following type properties can be set for marker formats in a chart.
These are shown in the same order as in the Excel format dialog.

    automatic
    none
    square
    diamond
    triangle
    x
    star
    short_dash
    long_dash
    circle
    plus

The automatic type is a special case which turns on a marker using the
default marker style for the particular series number.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :marker     => { :type => 'automatic' }
    )

If automatic is on then other marker properties such as size,
border or fill cannot be set.

The size property sets the size of the marker and is generally used
in conjunction with type.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :marker     => { :type => 'diamond', size => 7 }
    )

Nested border and fill properties can also be set for a marker.
See the [CHART FORMATTING][] section below.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :marker     => {
            :type    => 'square',
            :size    => 5,
            :border  => { :color => 'red' },
            :fill    => { :color => 'yellow' },
        }
    )

##### <a name="series_trendline" class="anchor" href="#series_trendline"><span class="octicon octicon-link" /></a>:trendline

A trendline can be added to a chart series to indicate trends in the data
such as a moving average or a polynomial fit.

The following properties can be set for trendlines in a chart series.

    :type
    :order       (for polynomial trends)
    :period      (for moving average)
    :forward     (for all except moving average)
    :backward    (for all except moving average)
    :name
    :line

The type property sets the type of trendline in the series.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :trendline  => { :type => 'linear' }
    )

The available trendline types are:

    exponential
    linear
    log
    moving_average
    polynomial
    power

A polynomial trendline can also specify the order of the polynomial.
The default value is 2.

    chart.add_series(
        :values    => '=Sheet1!$B$1:$B$5',
        :trendline => {
            :type  => 'polynomial',
            :order => 3
        }
    )

A moving_average trendline can also specify the period of the moving average.
The default value is 2.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :trendline  => {
            :type   => 'moving_average',
            :period => 3,
        }
    )

The forward and backward properties set the forecast period of the trendline.

    chart.add_series(
        :values    => '=Sheet1!$B$1:$B$5',
        :trendline => {
            :type     => 'linear',
            :forward  => 0.5,
            :backward => 0.5,
        }
    )

The name property sets an optional name for the trendline that will appear
in the chart legend.
If it isn't specified the Excel default name will be displayed.
This is usually a combination of the trendline type and the series name.

    chart.add_series(
        :values    => '=Sheet1!$B$1:$B$5',
        :trendline => {
            :type => 'linear',
            :name => 'Interpolated trend',
        }
    )

Several of these properties can be set in one go:

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :trendline  => {
            :type     => 'linear',
            :name     => 'My trend name',
            :forward  => 0.5,
            :backward => 0.5,
            :line     => {
                :color     => 'red',
                :width     => 1,
                :dash_type => 'long_dash',
            }
        }
    )

Trendlines cannot be added to series in a stacked chart or pie chart, radar
chart or (when implemented) to 3D, surface, or doughnut charts.

##### <a name="series_error_bars" class="anchor" href="#series_error_bars"><span class="octicon octicon-link" /></a>Error Bars

Error bars can be added to a chart series to indicate error bounds in the data.
The error bars can be vertical y_error_bars (the most common type) or
horizontal x_error_bars (for Bar and Scatter charts only).

The following properties can be set for error bars in a chart series.

    :type
    :value       (for all types except standard error)
    :direction
    :end_style
    :line

The type property sets the type of error bars in the series.

    chart.add_series(
        :values       => '=Sheet1!$B$1:$B$5',
        :y_error_bars => { :type => 'standard_error' },
    )

The available error bars types are available:

    fixed
    percentage
    standard_deviation
    standard_error

Note, the "custom" error bars type is not supported.

All error bar types, except for standard_error must also have a value associated
with it for the error bounds:

    chart.add_series(
        :values       => '=Sheet1!$B$1:$B$5',
        :y_error_bars => {
            :type  => 'percentage',
            :value => 5
        }
    )

The direction property sets the direction of the error bars. It should be one
of the following:

    plus    # Positive direction only.
    minus   # Negative direction only.
    both    # Plus and minus directions, The default.

The end_style property sets the style of the error bar end cap.
The options are 1 (the default) or 0 (for no end cap):

    chart.add_series(
        :values       => '=Sheet1!$B$1:$B$5',
        :y_error_bars => {
            :type      => 'fixed',
            :value     => 2,
            :end_style => 0,
            :direction => 'minus'
        },
    )

##### <a name="data_labels" class="anchor" href="#data_labels"><span class="octicon octicon-link" /></a>Data Labels

Data labels can be added to a chart series to indicate the values of the plotted
data points.

The following properties can be set for data_labels formats in a chart.

    :value
    :category
    :series_name
    :position
    :leader_lines
    :percentage

The value property turns on the Value data label for a series.

    chart.add_series(
        :values      => '=Sheet1!$B$1:$B$5',
        :data_labels => { :value => 1 }
    )

The category property turns on the Category Name data label for a series.

    chart.add_series(
        :values      => '=Sheet1!$B$1:$B$5',
        :data_labels => { :category => 1 }
    )

The series_name property turns on the Series Name data label for a series.

    chart.add_series(
        :values      => '=Sheet1!$B$1:$B$5',
        :data_labels => { :series_name => 1 }
    )

The position property is used to position the data label for a series.

    chart.add_series(
        :values      => '=Sheet1!$B$1:$B$5',
        :data_labels => { :value => 1, :position => 'center' }
    )

Valid positions are:

    center
    right
    left
    top
    bottom
    above           # Same as top
    below           # Same as bottom
    inside_end      # Pie chart mainly.
    outside_end     # Pie chart mainly.
    best_fit        # Pie chart mainly.

The percentage property is used to turn on the display of data labels as
a Percentage for a series. It is mainly used for pie charts.

    chart.add_series(
        :values      => '=Sheet1!$B$1:$B$5',
        :data_labels => { :percentage => 1 }
    )

The leader_lines property is used to turn on Leader Lines for the data label for a series. It is mainly used for pie charts.

    chart.add_series(
        :values      => '=Sheet1!$B$1:$B$5',
        :data_labels => { :value => 1, :leader_lines => 1 }
    )

Note: Even when leader lines are turned on they aren't automatically visible
in Excel or WriteXLSX.
Due to an Excel limitation (or design) leader lines only appear if the data
label is moved manually or if the data labels are very close and need to be
adjusted automatically.

##### <a name="points" class="anchor" href="#points"><span class="octicon octicon-link" /></a>Points

In general formatting is applied to an entire series in a chart. However,
it is occasionally required to format individual points in a series.
In particular this is required for Pie charts where each segment is represented
by a point.

In these cases it is possible to use the points property of `add_series()`:

    chart.add_series(
        :values => '=Sheet1!$A$1:$A$3',
        :points => [
            { :fill => { :color => '#FF0000' } },
            { :fill => { :color => '#CC0000' } },
            { :fill => { :color => '#990000' } },
        ]
    )

The points property takes an array ref of format options
(see the [CHART FORMATTING][] section below).
To assign default properties to points in a series pass `nil` values in the array:

    # Format point 3 of 3 only.
    chart.add_series(
        :values => '=Sheet1!$A$1:$A$3',
        :points => [
            nil,
            nil,
            { :fill => { :color => '#990000' } },
        ]
    )

    # Format the first point only.
    chart.add_series(
        :values => '=Sheet1!$A$1:$A$3',
        :points => [ { :fill => { :color => '#FF0000' } } ]
    )

##### <a name="smooth" class="anchor" href="#smooth"><span class="octicon octicon-link" /></a>:Smooth

The `:smooth` option is used to set the smooth property of a line series.
It is only applicable to the Line and Scatter chart types.

    chart.add_series( :values => '=Sheet1!$C$1:$C$5',
                      :smooth => 1 )

### <a name="chart_formatting" class="anchor" href="#chart_formatting"><span class="octicon octicon-link" /></a>CHART FORMATTING

The following chart formatting properties can be set for any chart object
that they apply to (and that are supported by WriteXLSX) such as chart lines,
column fill areas, plot area borders, markers, gridlines and other chart
elements documented above.

    :line
    :border
    :fill

Chart formatting properties are generally set using hash.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => { :color => 'blue' }
    )

In some cases the format properties can be nested. For example a marker may
contain border and fill sub-properties.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => { :color => 'blue' },
        :marker     => {
            :type    => 'square',
            :size    => 5,
            :border  => { :color => 'red' },
            :fill    => { :color => 'yellow' },
        }
    )

#### <a name="formatting_line" class="anchor" href="#formatting_line"><span class="octicon octicon-link" /></a>:line

The line format is used to specify properties of line objects that appear
in a chart such as a plotted line on a chart or a border.

The following properties can be set for line formats in a chart.

    :none
    :color
    :width
    :dash_type

The none property is uses to turn the line off (it is always on by default
except in Scatter charts). This is useful if you wish to plot a series
with markers but without a line.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => { :none => 1 }
    )

The color property sets the color of the line.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => { :color => 'red' }
    )

The available colours are shown in the main WriteXLSX documentation. It is also possible to set the colour of a line with a HTML style RGB colour:

    chart.add_series(
        :line       => { :color => '#FF0000' }
    )

The width property sets the width of the line. It should be specified in increments of 0.25 of a point as in Excel.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => { :width => 3.25 }
    )

The dash_type property sets the dash style of the line.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => { :dash_type => 'dash_dot' }
    )

The following dash_type values are available. They are shown in the order that they appear in the Excel dialog.

    solid
    round_dot
    square_dot
    dash
    dash_dot
    long_dash
    long_dash_dot
    long_dash_dot_dot

The default line style is solid.

More than one line property can be specified at a time:

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :line       => {
            :color     => 'red',
            :width     => 1.25,
            :dash_type => 'square_dot'
        }
    )

#### <a name="border_formatting" class="anchor" href="#border_formatting"><span class="octicon octicon-link" /></a>:border

The border property is a synonym for line.

It can be used as a descriptive substitute for line in chart types such as Bar
and Column that have a border and fill style rather than a line style.
In general chart objects with a border property will also have a fill property.

#### <a name="fill_formatting" class="anchor" href="#fill_formatting"><span class="octicon octicon-link" /></a>:fill

The fill format is used to specify filled areas of chart objects such as
the interior of a column or the background of the chart itself.

The following properties can be set for fill formats in a chart.

    none
    color

The none property is used to turn the fill property off (it is generally on by default).

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :fill       => { :none => 1 }
    )

The color property sets the colour of the fill area.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :fill       => { :color => 'red' }
    )

The available colours are shown in the main WriteXLSX documentation. It is also possible to set the colour of a fill with a HTML style RGB colour:

    chart.add_series(
        :fill       => { :color => '#FF0000' }
    )

The fill format is generally used in conjunction with a border format which has the same properties as a line format.

    chart.add_series(
        :values     => '=Sheet1!$B$1:$B$5',
        :border     => { :color => 'red' },
        :fill       => { :color => 'yellow' }
    )

#### <a name="values_and_category_axes" class="anchor" href="#value_and_category_axes"><span class="octicon octicon-link" /></a>Value and Category Axes

Excel differentiates between a chart axis that is used for series *categories*
and an axis that is used for series *values*.

Since Excel treats the axes differently it also handles their formatting
differently and exposes different properties for each.

As such some of WriteXLSX axis properties can be set for a value axis,
some can be set for a category axis and some properties can be set for both.

For example the min and max properties can only be set for value axes and reverse
can be set for both. The type of axis that a property applies to is shown in
the [set_x_axis()][] section of the documentation above.

Some charts such as `Scatter` and `Stock` have two value axes.

[CHART FONTS]: chart_fonts.html#chart_fonts
[CHART FORMATTING]: chart.html#chart_formatting
[SERIES OPTIONS]: chart.html#series_options
[insert_chart()]: worksheet.html#insert_chart
[set_x_axis()]: #set_x_axis
