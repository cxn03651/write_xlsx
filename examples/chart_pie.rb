#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A demo of a Pie chart in Excel::Writer::XLSX.
#
# The demo also shows how to set segment colours. It is possible to
# define chart colors for most types of WrtieXLSX charts
# via the add_series() method. However, Pie charts are a special case
# since each segment is represented as a point so it is necessary to
# assign formatting to each point in the series.
#
# reverse(c), March 2011, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('chart_pie.xlsx')
worksheet = workbook.add_worksheet
bold      = workbook.add_format(bold: 1)

# Add the worksheet data that the charts will refer to.
headings = %w[Category Values]
data = [
  %w[Apple Cherry Pecan],
  [60,       30,       10]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)

# Create a new chart object. In this case an embedded chart.
chart1 = workbook.add_chart(type: 'pie', embedded: 1)

# Configure the series. Note the use of the array ref to define ranges:
# [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
# See below for an alternative syntax.
chart1.add_series(
  name:       'Pie sales data',
  categories: ['Sheet1', 1, 3, 0, 0],
  values:     ['Sheet1', 1, 3, 1, 1]
)

# Add a title.
chart1.set_title(name: 'Popular Pie Types')

# Set an Excel chart style. Blue colors with white outline and shadow.
chart1.set_style(10)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'C2', chart1,
  x_offset: 25, y_offset: 10
)

#
# Create a Pie chart with user defined segment colors.
#

# Create an example Pie chart like above.
chart2 = workbook.add_chart(type: 'pie', embedded: 1)

# Configure the series.
chart2.add_series(
  name:       'Pie sales data',
  categories: '=Sheet1!$A$2:$A$4',
  values:     '=Sheet1!$B$2:$B$4',
  points:     [
    { fill: { color: '#5ABA10' } },
    { fill: { color: '#FE110E' } },
    { fill: { color: '#CA5C05' } }
  ]
)

# Add a title.
chart2.set_title(name: 'Pie Chart with user defined colors')

worksheet.insert_chart(
  'C18', chart2,
  x_offset: 25, y_offset: 10
)

workbook.close
