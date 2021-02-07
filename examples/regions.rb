#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# An example of how to use the WriteXLSX gem to write a basic
# Excel workbook with multiple worksheets.
#
# reverse(c), March 2001, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('regions.xlsx')

# Add some worksheets
north = workbook.add_worksheet('North')
south = workbook.add_worksheet('South')
east  = workbook.add_worksheet('East')
west  = workbook.add_worksheet('West')

# Add a Format
format = workbook.add_format
format.set_bold
format.set_color('blue')

# Add a caption to each worksheet
workbook.sheets.each do |worksheet|
  worksheet.write(0, 0, 'Sales', format)
end

# Write some data
north.write(0, 1, 200000)
south.write(0, 1, 100000)
east.write(0, 1, 150000)
west.write(0, 1, 100000)

# Set the active worksheet
south.activate

# Set the width of the first column
south.set_column(0, 0, 20)

# Set the active cell
south.set_selection(0, 1)

workbook.close
