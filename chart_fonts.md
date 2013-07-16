---
layout: default
title: Chart Fonts
---
### <a name="chart_fonts" class="anchor" href="#chart_fonts"><span class="octicon octicon-link" /></a>CHART FONTS

The following font properties can be set for any chart object that they apply to
(and that are supported by WriteXLSX) such as chart titles, axis labels
and axis numbering.
They correspond to the equivalent Worksheet cell Format object properties.
See ["FORMAT_METHODS"](format.html#format) for more information.

    :name
    :size
    :bold
    :italic
    :underline
    :rotation
    :color

The following explains the available font properties:

##### <a name="name" class="anchor" href="#name"><span class="octicon octicon-link" /></a>:name
Set the font name:

    chart.set_x_axis( :num_font => { :name => 'Arial' } )

##### <a name="size" class="anchor" href="#size"><span class="octicon octicon-link" /></a>:size
Set the font size:

    chart.set_x_axis( :num_font => { :name => 'Arial', :size => 10 } )

##### <a name="bold" class="anchor" href="#bold"><span class="octicon octicon-link" /></a>:bold
Set the font bold property, should be 0 or 1:

    chart.set_x_axis( :num_font => { :bold => 1 } )

##### <a name="italic" class="anchor" href="#italic"><span class="octicon octicon-link" /></a>:italic
Set the font italic property, should be 0 or 1:

    chart.set_x_axis( :num_font => { :italic => 1 } )

##### <a name="underline" class="anchor" href="#underline"><span class="octicon octicon-link" /></a>:underline
Set the font underline property, should be 0 or 1:

    chart.set_x_axis( :num_font => { :underline => 1 } )

##### <a name="rotation" class="anchor" href="#rotation"><span class="octicon octicon-link" /></a>:rotation
Set the font rotation in the range -90 to 90:

    chart.set_x_axis( :num_font => { :rotation => 45 } )

This is useful for displaying large axis data such as dates in a more compact format.

##### <a name="color" class="anchor" href="#color"><span class="octicon octicon-link" /></a>:color
Set the font color property. Can be a color index, a color name or HTML style RGB colour:

    chart.set_x_axis( :num_font => { :color => 'red' } )
    chart.set_y_axis( :num_font => { :color => '#92D050' } )

Here is an example of Font formatting in a Chart program:

    # Format the chart title.
    chart.set_title(
        :name      => 'Sales Results Chart',
        :name_font => {
            :name  => 'Calibri',
            :color => 'yellow',
        }
    )

    # Format the X-axis.
    chart.set_x_axis(
        :name      => 'Month',
        :name_font => {
            :name  => 'Arial',
            :color => '#92D050'
        },
        :num_font => {
            :name  => 'Courier New',
            :color => '#00B0F0',
        }
    )

    # Format the Y-axis.
    chart.set_y_axis(
        :name      => 'Sales (1000 units)',
        :name_font => {
            :name      => 'Century',
            :underline => 1,
            :color     => 'red'
        },
        :num_font => {
            :bold   => 1,
            :italic => 1,
            :color  => '#7030A0',
        }
    )
