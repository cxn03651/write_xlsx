#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# A demo of a Pie chart in Excel::Writer::XLSX.
#
# reverse(c), March 2011, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, cxn03651@msj.biglobe.ne.jp
#

require 'rubygems'
require 'write_xlsx'

workbook  = WriteXLSX.new('chart_pie.xlsx')
worksheet = workbook.add_worksheet
bold      = workbook.add_format(:bold => 1)

# Add the worksheet data that the charts will refer to.
headings = [ 'Category', 'Values' ]
data = [
    [ 'Apple', 'Cherry', 'Pecan' ],
    [ 60,       30,       10     ]
]

worksheet.write('A1', headings, bold)
worksheet.write('A2', data)

# Create a new chart object. In this case an embedded chart.
chart = workbook.add_chart(:type => 'pie', :embedded => 1)

# Configure the series. Note the use of the array ref to define ranges:
# [ $sheetname, $row_start, $row_end, $col_start, $col_end ].
chart.add_series(
    :name       => 'Pie sales data',
    :categories => [ 'Sheet1', 1, 3, 0, 0 ],
    :values     => [ 'Sheet1', 1, 3, 1, 1 ]
)

# Add a title.
chart.set_title(:name => 'Popular Pie Types')

# Set an Excel chart style. Blue colors with white outline and shadow.
chart.set_style(10)

# Insert the chart into the worksheet (with an offset).
worksheet.insert_chart('C2', chart, 25, 10)

workbook.close
