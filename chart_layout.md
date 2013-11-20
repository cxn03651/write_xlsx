---
layout: default
title: Chart Layout
---

### <a name="chart_layout" class="anchor" href="#chart_layout"><span class="octicon octicon-link" /></a>CHART LAYOUT

The position of the chart in the worksheet is controlled by the set_size method shown above.

It is also possible to change the layout of the following chart sub-objects:

    :plotarea
    :legend
    :title
    :x_axis caption
    :y_axis caption

Here are some examples:

    chart.set_plotarea(
      :layout => {
        :x      => 0.35,
        :y      => 0.26,
        :width  => 0.62,
        :height => 0.50
      }
    )

    chart.set_legend(
      :layout => {
        :x      => 0.80,
        :y      => 0.37,
        :width  => 0.12,
        :height => 0.25
      }
    )

    chartset_title(
      :name   => 'Title',
      :layout => {
        :x      => 0.80,
        :y      => 0.37,
        :width  => 0.12,
        :height => 0.25
      }
    )

    chartset_x_axis(
      :name        => 'X axis,
      :name_layout => {
        :x      => 0.80,
        :y      => 0.37
      }
    )

Note that it is only possible to change the width and height for the plotarea
and legend objects. For the other text based objects the width and height are
chaged by the font dimensions.

The layout units must be a float in the range 0 < x <= 1 and are expressed
as a percentage of the chart dimensions as shown below:

![Chart object layout.](images/examples/layout.png)

From this the layout units are calculated as follows:

    layout:
      width  = w / W
      height = h / H
      x      = a / W
      y      = b / H

These units area slightly cumbersome but are required by Excel so that the chart object positions
remain relative to each other if the cahrt is resized by the user.

Note that for plotarea the origin is the top left corner in the plotarea itself
and does not take into account the axes.
