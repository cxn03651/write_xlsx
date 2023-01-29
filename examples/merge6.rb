#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX merge_cells workbook
# method with Unicode strings.
#
# reverse(c), December 2005, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Create a new workbook and add a worksheet
workbook  = WriteXLSX.new('merge6.xlsx')
worksheet = workbook.add_worksheet

# Increase the cell size of the merged cells to highlight the formatting.
(2..9).each { |row| worksheet.set_row(row, 36) }
worksheet.set_column('B:D', 25)

# Format for the merged cells.
format = workbook.add_format(
  border: 6,
  bold:   1,
  color:  'red',
  size:   20,
  valign: 'vcentre',
  align:  'left',
  indent: 1
)

###############################################################################
#
# Write an Ascii string.
#
worksheet.merge_range('B3:D4', 'ASCII: A simple string', format)

###############################################################################
#
# Write a UTF-8 Unicode string.
#
smiley = '☺'
worksheet.merge_range('B6:D7', "UTF-8: A Unicode smiley #{smiley}", format)

workbook.close
