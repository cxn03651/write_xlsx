---
layout: default
title: Sparklines
---
### <a name="sparklines" class="anchor" href="#sparklines"><span class="octicon octicon-link" /></a>SPARKLINES IN EXCEL

Sparklines are a feature of Excel 2010+ which allows you to add small charts
to worksheet cells.
These are useful for showing visual trends in data in a compact format.

In WriteXLSX Sparklines can be added to cells using the
[add_sparkline()][] worksheet method:

    worksheet.add_sparkline(
        {
            :location => 'F2',
            :range    => 'Sheet1!A2:E2',
            :type     => 'column',
            :style    => 12,
        }
    )

![Sparklines example.](images/sparklines1.jpg)

Note: Sparklines are a feature of Excel 2010+ only.
You can write them to an XLSX file that can be read by Excel 2007
but they won't be displayed.

#### <a name="add_sparkline" class="anchor" href="#add_sparkline"><span class="octicon octicon-link" /></a>add_sparkline( { parameter => 'value', ... } )

The [add_sparkline()][] worksheet method is used to add sparklines to a cell
or a range of cells.

The parameters to [add_sparkline()][] must be passed in a hash.
The main sparkline parameters are:

    :location        (required)
    :range           (required)
    :type
    :style

    :markers
    :negative_points
    :axis
    :reverse

Other, less commonly used parameters are:

    :high_point
    :low_point
    :first_point
    :last_point
    :max
    :min
    :empty_cells
    :show_hidden
    :date_axis
    :weight

    :series_color
    :negative_color
    :markers_color
    :first_color
    :last_color
    :high_color
    :low_color

These parameters are explained in the sections below:

##### <a name="location" class="anchor" href="#location"><span class="octicon octicon-link" /></a>:location

This is the cell where the sparkline will be displayed:

    :location => 'F1'

The `:location` should be a single cell.
(For multiple cells see [Grouped Sparklines] below).

#### <a name="range" class="anchor" href="#range"><span class="octicon octicon-link" /></a>:range

This specifies the cell data range that the sparkline will plot:

    worksheet.add_sparkline(
        {
            :location => 'F1',
            :range    => 'A1:E1',
        }
    )

The `:range` should be a 2D array.
(For 3D arrays of cells see [Grouped Sparklines] below).

If `:range` is not on the same worksheet you can specify its location
using the usual Excel notation:

            :range => 'Sheet1!A1:E1',

If the worksheet contains spaces or special characters you should quote
the worksheet name in the same way that Excel does:

            :range => q('Monthly Data'!A1:E1),

##### <a name="type" class="anchor" href="#type"><span class="octicon octicon-link" /></a>:type

Specifies the type of sparkline. There are 3 available sparkline types:

    line    (default)
    column
    win_loss

For example:

    {
        :location => 'F1',
        :range    => 'A1:E1',
        :type     => 'column',
    }

##### <a name="style" class="anchor" href="#style"><span class="octicon octicon-link" /></a>:style

Excel provides 36 built-in Sparkline styles in 6 groups of 6.
The `:style` parameter can be used to replicate these and should be
a corresponding number from 1 .. 36.

    {
        :location => 'A14',
        :range    => 'Sheet2!A2:J2',
        :style    => 3,
    }

The `:style` number starts in the top left of the style grid and runs left
to right. The default style is 1. It is possible to override colour elements
of the sparklines using the \*_color parameters below.

##### <a name="markers" class="anchor" href="#markers"><span class="octicon octicon-link" /></a>:markers

Turn on the markers for line style sparklines.

    {
        :location => 'A6',
        :range    => 'Sheet2!A1:J1',
        :markers  => 1,
    }

Markers aren't shown in Excel for column and win_loss sparklines.

##### <a name="negative_points" class="anchor" href="#negative_points"><span class="octicon octicon-link" /></a>:negative_points

Highlight negative values in a sparkline range.
This is usually required with win_loss sparklines.

    {
        :location        => 'A21',
        :range           => 'Sheet2!A3:J3',
        :type            => 'win_loss',
        :negative_points => 1,
    }

##### <a name="axis" class="anchor" href="#axis"><span class="octicon octicon-link" /></a>:axis

Display a horizontal axis in the sparkline:

    {
        :location => 'A10',
        :range    => 'Sheet2!A1:J1',
        :axis     => 1,
    }

##### <a name="reverse" class="anchor" href="#reverse"><span class="octicon octicon-link" /></a>:reverse

Plot the data from right-to-left instead of the default left-to-right:

    {
        :location => 'A24',
        :range    => 'Sheet2!A4:J4',
        :type     => 'column',
        :reverse  => 1,
    }

##### <a name="weight" class="anchor" href="#weight"><span class="octicon octicon-link" /></a>:weight

Adjust the default line weight (thickness) for line style sparklines.

     :weight => 0.25,

The weight value should be one of the following values allowed by Excel:

    0.25  0.5   0.75
    1     1.25
    2.25
    3
    4.25
    6

##### <a name="high_low_first_last_point" class="anchor" href="#high_low_first_last_point"><span class="octicon octicon-link" /></a>high_point, low_point, first_point, last_point

Highlight points in a sparkline range.

        :high_point  => 1,
        :low_point   => 1,
        :first_point => 1,
        :last_point  => 1,

##### <a name="max_min" class="anchor" href="#max_min"><span class="octicon octicon-link" /></a>:max, :min

Specify the maximum and minimum vertical axis values:

        :max         => 0.5,
        :min         => -0.5,

As a special case you can set the maximum and minimum to be for a group of sparklines rather than one:

        :max         => 'group',

See [Grouped Sparklines][] below.

##### <a name="empty_cells" class="anchor" href="#empty_cells"><span class="octicon octicon-link" /></a>:empty_cells

Define how empty cells are handled in a sparkline.

    :empty_cells => 'zero',

The available options are:

    gaps   : show empty cells as gaps (the default).
    zero   : plot empty cells as 0.
    connect: Connect points with a line ("line" type  sparklines only).

##### <a name="show_hidden" class="anchor" href="#show_hidden"><span class="octicon octicon-link" /></a>:show_hidden

Plot data in hidden rows and columns:

    :show_hidden => 1,

Note, this option is off by default.

##### <a name="data_axis" class="anchor" href="#data_axis"><span class="octicon octicon-link" /></a>:date_axis

Specify an alternative date axis for the sparkline.
This is useful if the data being plotted isn't at fixed width intervals:

    {
        :location  => 'F3',
        :range     => 'A3:E3',
        :date_axis => 'A4:E4',
    }

The number of cells in the date range should correspond to the number
of cells in the data range.

##### <a name="series_color" class="anchor" href="#series_color"><span class="octicon octicon-link" /></a>:series_color

It is possible to override the colour of a sparkline style using the following parameters:

    :series_color
    :negative_color
    :markers_color
    :first_color
    :last_color
    :high_color
    :low_color

The color should be specified as a HTML style #rrggbb hex value:

    {
        :location     => 'A18',
        :range        => 'Sheet2!A2:J2',
        :type         => 'column',
        :series_color => '#E965E0',
    }

#### <a name="grouped_sparklines" class="anchor" href="#grouped_sparklines"><span class="octicon octicon-link" /></a>Grouped Sparklines

The `add_sparkline()` worksheet method can be used multiple times to write
as many sparklines as are required in a worksheet.

However, it is sometimes necessary to group contiguous sparklines so that
changes that are applied to one are applied to all.
In Excel this is achieved by selecting a 3D range of cells for the data range
and a 2D range of cells for the location.

In WriteXLSX, you can simulate this by passing an array of values to location
and range:

    {
        :location => [ 'A27',          'A28',          'A29'          ],
        :range    => [ 'Sheet2!A5:J5', 'Sheet2!A6:J6', 'Sheet2!A7:J7' ],
        :markers  => 1,
    }

#### <a name="sparkline_examples" class="anchor" href="#sparkline_examples"><span class="octicon octicon-link" /></a>Sparkline examples

See the
[sparklines1.rb](examples.html#sparklines1)
and
[sparklines2.rb](examples.html#sparklines2)
example programs in the examples directory of the distro.


[Grouped Sparklines]: sparklines.html#grouped_sparklines
[add_sparkline()]: worksheet.html#add_sparkline
