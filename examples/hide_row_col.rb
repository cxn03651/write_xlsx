#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# Example of how to hide rows and columns in Excel::Writer::XLSX. In order to
# hide rows without setting each one, (of approximately 1 million rows),
# Excel uses an optimisation to hide all rows that don't have data.
#
# reverse ('(c)'), December 2012, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook  = WriteXLSX.new('hide_row_col.xlsx')
worksheet = workbook.add_worksheet

# Write some data
worksheet.write('D1', 'Some hidden columns.')
worksheet.write('A8', 'Some hidden rows.')

# Hide all rows without data.
worksheet.set_default_row(nil, 1)

# Set emptys row that we do want to display. All other will be hidden.
(1..6).each { |row| worksheet.set_row(row, 15) }

# Hide a range of columns.
worksheet.set_column('G:XFD', nil, nil, 1)

workbook.close
