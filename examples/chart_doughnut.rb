#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A demo of a Doughnut chart in Excel::Writer::XLSX.
#
# The demo also shows how to set segment colours. It is possible to define
# chart colors for most types of Excel::Writer::XLSX charts via the
# add_series() method. However, Pie and Doughtnut charts are a special case
# since each segment is represented as a point so it is necessary to assign
# formatting to each point in the series.
#
# reverse ('(c)'), March 2011, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('chart_doughnut.xlsx')
worksheet = workbook.add_worksheet
bold      = workbook.add_format(bold: 1)

# Add the worksheet data that the charts will refer to.
headings = %w[Category Values]
data = [
  %w[Glazed Chocolate Cream],
  [50,       35,          15]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)

# Create a new chart object. In this case an embedded chart.
chart1 = workbook.add_chart(type: 'doughnut', embedded: 1)

# Configure the series. Note the use of the array ref to define ranges:
# [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
# See below for an alternative syntax.
chart1.add_series(
  name:       'Doughnut sales data',
  categories: ['Sheet1', 1, 3, 0, 0],
  values:     ['Sheet1', 1, 3, 1, 1]
)

# Add a title.
chart1.set_title(name: 'Popular Doughnut Types')

# Set an Excel chart style. Colors with white outline and shadow.
chart1.set_style(10)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'C2', chart1,
  x_offset: 25, y_offset: 10
)

#
# Create a Doughnut chart with user defined segment colors.
#

# Create an example Doughnut chart like above.
chart2 = workbook.add_chart(type: 'doughnut', embedded: 1)

# Configure the series and add user defined segment colours.
chart2.add_series(
  name:       'Doughnut sales data',
  categories: '=Sheet1!$A$2:$A$4',
  values:     '=Sheet1!$B$2:$B$4',
  points:     [
    { fill: { color: '#FA58D0' } },
    { fill: { color: '#61210B' } },
    { fill: { color: '#F5F6CE' } }
  ]
)

# Add a title.
chart2.set_title(name: 'Doughnut Chart with user defined colors')

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'C18', chart2,
  x_offset: 25, y_offset: 10
)

#
# Create a Doughnut chart with rotation of the segments.
#

# Create an example Doughnut chart like above.
chart3 = workbook.add_chart(type: 'doughnut', embedded: 1)

# Configure the series.
chart3.add_series(
  name:       'Doughnut sales data',
  categories: '=Sheet1!$A$2:$A$4',
  values:     '=Sheet1!$B$2:$B$4'
)

# Add a title.
chart3.set_title(name: 'Doughnut Chart with segment rotation')

# Change the angle/rotation of the first segment.
chart3.set_rotation(90)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'C34', chart3,
  x_offset: 25, y_offset: 10
)

#
# Create a Doughnut chart with user defined hole size.
#

# Create an example Doughnut chart like above.
chart4 = workbook.add_chart(type: 'doughnut', embedded: 1)

# Configure the series.
chart4.add_series(
  name:       'Doughnut sales data',
  categories: '=Sheet1!$A$2:$A$4',
  values:     '=Sheet1!$B$2:$B$4'
)

# Add a title.
chart4.set_title(name: 'Doughnut Chart with user defined hole size')

# Change the hole size.
chart4.set_hole_size(33)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'C50', chart4,
  x_offset: 25, y_offset: 10
)

workbook.close
