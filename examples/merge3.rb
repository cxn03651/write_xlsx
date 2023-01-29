#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

###############################################################################
#
# Example of how to use Excel::Writer::XLSX to write a hyperlink in a
# merged cell.
#
# reverse(c), September 2002, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Create a new workbook and add a worksheet
workbook  = WriteXLSX.new('merge3.xlsx')
worksheet = workbook.add_worksheet

# Increase the cell size of the merged cells to highlight the formatting.
[3, 6, 7].each { |row| worksheet.set_row(row, 30) }
worksheet.set_column('B:D', 20)

###############################################################################
#
# Example: Merge cells containing a hyperlink using merge_range().
#
format = workbook.add_format(
  border:    1,
  underline: 1,
  color:     'blue',
  align:     'center',
  valign:    'vcenter'
)

# Merge 3 cells
worksheet.merge_range('B4:D4', 'http://www.perl.com', format)

# Merge 3 cells over two rows
worksheet.merge_range('B7:D8', 'http://www.perl.com', format)

workbook.close
