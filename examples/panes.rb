#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# Example of using the Excel::Writer::XLSX module to create worksheet panes.
#
# reverse(c), May 2001, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Create a new workbook called simple.xls and add a worksheet
workbook  = WriteXLSX.new('panes.xlsx')

worksheet1 = workbook.add_worksheet('Panes 1')
worksheet2 = workbook.add_worksheet('Panes 2')
worksheet3 = workbook.add_worksheet('Panes 3')
worksheet4 = workbook.add_worksheet('Panes 4')

# Freeze panes
worksheet1.freeze_panes(1, 0)    # 1 row

worksheet2.freeze_panes(0, 1)    # 1 column
worksheet3.freeze_panes(1, 1)    # 1 row and column

# Split panes.
# The divisions must be specified in terms of row and column dimensions.
# The default row height is 15 and the default column width is 8.43
#
worksheet4.split_panes(15, 8.43)    # 1 row and column

#######################################################################
#
# Set up some formatting and text to highlight the panes
#

header = workbook.add_format(
  align:    'center',
  valign:   'vcenter',
  fg_color: '#C3FFC0'
)

center = workbook.add_format(align: 'center')

#######################################################################
#
# Sheet 1
#

worksheet1.set_column('A:I', 16)
worksheet1.set_row(0, 20)
worksheet1.set_selection('C3')

9.times { |i| worksheet1.write(0, i, 'Scroll down', header) }
(1..100).each do |i|
  9.times { |j| worksheet1.write(i, j, i + 1, center) }
end

#######################################################################
#
# Sheet 2
#

worksheet2.set_column('A:A', 16)
worksheet2.set_selection('C3')

50.times do |i|
  worksheet2.set_row(i, 15)
  worksheet2.write(i, 0, 'Scroll right', header)
end

50.times do |i|
  (1..25).each { |j| worksheet2.write(i, j, j, center) }
end

#######################################################################
#
# Sheet 3
#

worksheet3.set_column('A:Z', 16)
worksheet3.set_selection('C3')

worksheet3.write(0, 0, '', header)

(1..25).each { |i| worksheet3.write(0, i, 'Scroll down', header) }
(1..49).each { |i| worksheet3.write(i, 0, 'Scroll right', header) }
(1..49).each do |i|
  (1..25).each { |j| worksheet3.write(i, j, j, center) }
end

#######################################################################
#
# Sheet 4
#

worksheet4.set_selection('C3')

(1..25).each { |i| worksheet4.write(0, i, 'Scroll', center) }
(1..49).each { |i| worksheet4.write(i, 0, 'Scroll', center) }
(1..49).each do |i|
  (1..25).each { |j| worksheet4.write(i, j, j, center) }
end

workbook.close
