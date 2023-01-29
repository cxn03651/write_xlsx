#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A demo of an Area chart in Excel::Writer::XLSX.
#
# reverse('ｩ'), March 2011, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('chart_column.xlsx')
worksheet = workbook.add_worksheet
bold      = workbook.add_format(bold: 1)

# Add the worksheet data that the charts will refer to.
headings = ['Number', 'Batch 1', 'Batch 2']
data = [
  [2, 3, 4, 5, 6, 7],
  [10, 40, 50, 20, 10, 50],
  [30, 60, 70, 50, 40, 30]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)

# Create a new chart object. In this case an embedded chart.
chart = workbook.add_chart(type: 'column', embedded: 1)

# Configure the first series.
chart.add_series(
  name:       '=Sheet1!$B$1',
  categories: '=Sheet1!$A$2:$A$7',
  values:     '=Sheet1!$B$2:$B$7'
)

# Configure second series. Note alternative use of array ref to define
# ranges: [ sheetname, row_start, row_end, col_start, col_end ].
chart.add_series(
  name:       '=Sheet1!$C$1',
  categories: ['Sheet1', 1, 6, 0, 0],
  values:     ['Sheet1', 1, 6, 2, 2]
)

# Add a chart title and some axis labels.
chart.set_title(name: 'Results of sample analysis')
chart.set_x_axis(name: 'Test number')
chart.set_y_axis(name: 'Sample length (mm)')

# Set an Excel chart style. Blue colors with white outline and shadow.
chart.set_style(11)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'D2', chart,
  x_offset: 25, y_offset: 10
)

workbook.close
