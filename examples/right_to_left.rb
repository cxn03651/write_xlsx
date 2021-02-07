#!/usr/bin/env ruby

#######################################################################
#
# Example of how to change the default worksheet direction from
# left-to-right to right-to-left as required by some eastern verions
# of Excel.
#
# reverse(c), January 2006, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

require 'write_xlsx'

workbook   = WriteXLSX.new('right_to_left.xlsx')
worksheet1 = workbook.add_worksheet
worksheet2 = workbook.add_worksheet

worksheet2.right_to_left

worksheet1.write(0, 0, 'Hello')    #  A1, B1, C1, ...
worksheet2.write(0, 0, 'Hello')    # ..., C1, B1, A1
workbook.close
