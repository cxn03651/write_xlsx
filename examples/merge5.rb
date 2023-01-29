#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

###############################################################################
#
# Example of how to use the Excel::Writer::XLSX merge_cells workbook
# method with complex formatting and rotation.
#
#
# reverse(c), September 2002, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Create a new workbook and add a worksheet
workbook  = WriteXLSX.new('merge5.xlsx')
worksheet = workbook.add_worksheet

# Increase the cell size of the merged cells to highlight the formatting.
(3..8).each { |row|  worksheet.set_row(row, 36) }
[1, 3, 5].each { |col| worksheet.set_column(col, col, 15) }

###############################################################################
#
# Rotation 1, letters run from top to bottom
#
format1 = workbook.add_format(
  border:   6,
  bold:     1,
  color:    'red',
  valign:   'vcentre',
  align:    'centre',
  rotation: 270
)

worksheet.merge_range('B4:B9', 'Rotation 270', format1)

###############################################################################
#
# Rotation 2, 90ｰ anticlockwise
#
format2 = workbook.add_format(
  border:   6,
  bold:     1,
  color:    'red',
  valign:   'vcentre',
  align:    'centre',
  rotation: 90
)

worksheet.merge_range('D4:D9', 'Rotation 90', format2)

###############################################################################
#
# Rotation 3, 90ｰ clockwise
#
format3 = workbook.add_format(
  border:   6,
  bold:     1,
  color:    'red',
  valign:   'vcentre',
  align:    'centre',
  rotation: -90
)

worksheet.merge_range('F4:F9', 'Rotation -90', format3)

workbook.close
