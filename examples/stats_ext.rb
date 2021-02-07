#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# Example of how to use functions with the WriteXLSX gem.
#
# This is a simple example of how to use functions that reference cells in
# other worksheets within the same workbook.
#
# reverse(c), March 2001, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

# Create a new workbook and add a worksheet
workbook   = WriteXLSX.new('stats_ext.xlsx')
worksheet1 = workbook.add_worksheet('Test results')
worksheet2 = workbook.add_worksheet('Data')

# Set the column width for column 1
worksheet1.set_column(0, 0, 20)

# Create a format for the headings
headings = workbook.add_format
headings.set_bold

# Create a numerical format
numformat = workbook.add_format
numformat.set_num_format('0.00')

# Write some statistical functions
worksheet1.write(0, 0, 'Count', headings)
worksheet1.write(0, 1, '=COUNT(Data!B2:B9)')

worksheet1.write(1, 0, 'Sum', headings)
worksheet1.write(1, 1, '=SUM(Data!B2:B9)')

worksheet1.write(2, 0, 'Average', headings)
worksheet1.write(2, 1, '=AVERAGE(Data!B2:B9)')

worksheet1.write(3, 0, 'Min', headings)
worksheet1.write(3, 1, '=MIN(Data!B2:B9)')

worksheet1.write(4, 0, 'Max', headings)
worksheet1.write(4, 1, '=MAX(Data!B2:B9)')

worksheet1.write(5, 0, 'Standard Deviation', headings)
worksheet1.write(5, 1, '=STDEV(Data!B2:B9)')

worksheet1.write(6, 0, 'Kurtosis', headings)
worksheet1.write(6, 1, '=KURT(Data!B2:B9)')

# Write the sample data
worksheet2.write(0, 0, 'Sample', headings)
worksheet2.write(1, 0, 1)
worksheet2.write(2, 0, 2)
worksheet2.write(3, 0, 3)
worksheet2.write(4, 0, 4)
worksheet2.write(5, 0, 5)
worksheet2.write(6, 0, 6)
worksheet2.write(7, 0, 7)
worksheet2.write(8, 0, 8)

worksheet2.write(0, 1, 'Length', headings)
worksheet2.write(1, 1, 25.4, numformat)
worksheet2.write(2, 1, 25.4, numformat)
worksheet2.write(3, 1, 24.8, numformat)
worksheet2.write(4, 1, 25.0, numformat)
worksheet2.write(5, 1, 25.3, numformat)
worksheet2.write(6, 1, 24.9, numformat)
worksheet2.write(7, 1, 25.2, numformat)
worksheet2.write(8, 1, 24.8, numformat)

workbook.close
