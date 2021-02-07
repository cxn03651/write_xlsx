#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

#######################################################################
#
# Example of how to hide a worksheet with Excel::Writer::XLSX.
#
# reverse('c'), April 2005, John McNamara, jmcnamara@cpan.org
# convert to ruby by Hideo, Nakamura, nakamura.hideo@gmail.com
#
require 'write_xlsx'

workbook   = WriteXLSX.new('hide_sheet.xlsx')
worksheet1 = workbook.add_worksheet
worksheet2 = workbook.add_worksheet
worksheet3 = workbook.add_worksheet

worksheet1.set_column('A:A', 30)
worksheet2.set_column('A:A', 30)
worksheet3.set_column('A:A', 30)

# Sheet2 won't be visible until it is unhidden in Excel.
worksheet2.hide

worksheet1.write(0, 0, 'Sheet2 is hidden')
worksheet2.write(0, 0, "Now it's my turn to find you.")
worksheet3.write(0, 0, 'Sheet2 is hidden')

workbook.close
