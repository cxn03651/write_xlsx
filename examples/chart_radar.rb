#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A demo of an Area chart in Excel::Writer::XLSX.
#
# reverse ('(c)'), October 2012, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('chart_radar.xlsx')
worksheet = workbook.add_worksheet
bold      = workbook.add_format(bold: 1)

# Add the worksheet data that the charts will refer to.
headings = ['Number', 'Batch 1', 'Batch 2']
data = [
  [2, 3, 4, 5, 6, 7],
  [30, 60, 70, 50, 40, 30],
  [25, 40, 50, 30, 50, 40]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)

# Create a new chart object. In this case an embedded chart.
chart1 = workbook.add_chart(type: 'radar', embedded: 1)

# Configure the first series.
chart1.add_series(
  name:       '=Sheet1!$B$1',
  categories: '=Sheet1!$A$2:$A$7',
  values:     '=Sheet1!$B$2:$B$7'
)

# Configure second series. Note alternative use of array ref to define
# ranges: [ sheetname, row_start, row_end, col_start, col_end ].
chart1.add_series(
  name:       '=Sheet1!$C$1',
  categories: ['Sheet1', 1, 6, 0, 0],
  values:     ['Sheet1', 1, 6, 2, 2]
)

# Add a chart title.
chart1.set_title(name: 'Results of sample analysis')

# Set an Excel chart style. Blue colors with white outline and shadow.
chart1.set_style(11)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'D2', chart1,
  x_offset: 25, y_offset: 10
)

#
# Create a with_markers chart sub-type
#
chart2 = workbook.add_chart(
  type:     'radar',
  embedded: 1,
  subtype:  'with_markers'
)

# Configure the first series.
chart2.add_series(
  name:       '=Sheet1!$B$1',
  categories: '=Sheet1!$A$2:$A$7',
  values:     '=Sheet1!$B$2:$B$7'
)

# Configure second series.
chart2.add_series(
  name:       '=Sheet1!$C$1',
  categories: ['Sheet1', 1, 6, 0, 0],
  values:     ['Sheet1', 1, 6, 2, 2]
)

# Add a chart title.
chart2.set_title(name: 'Stacked Chart')

# Set an Excel chart style. Blue colors with white outline and shadow.
chart2.set_style(12)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'D18', chart2,
  x_offset: 25, y_offset: 10
)

#
# Create a filled chart sub-type
#
chart3 = workbook.add_chart(
  type:     'radar',
  embedded: 1,
  subtype:  'filled'
)

# Configure the first series.
chart3.add_series(
  name:       '=Sheet1!$B$1',
  categories: '=Sheet1!$A$2:$A$7',
  values:     '=Sheet1!$B$2:$B$7'
)

# Configure second series.
chart3.add_series(
  name:       '=Sheet1!$C$1',
  categories: ['Sheet1', 1, 6, 0, 0],
  values:     ['Sheet1', 1, 6, 2, 2]
)

# Add a chart title.
chart3.set_title(name: 'Percent Stacked Chart')

# Set an Excel chart style. Blue colors with white outline and shadow.
chart3.set_style(13)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart(
  'D34', chart3,
  x_offset: 25, y_offset: 10
)

workbook.close
