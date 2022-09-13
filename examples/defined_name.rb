#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'write_xlsx'

workbook   = WriteXLSX.new('defined_name.xlsx')
worksheet1 = workbook.add_worksheet
worksheet2 = workbook.add_worksheet

# Define some global/workbook names.
workbook.define_name('Exchange_rate', '=0.96')
workbook.define_name('Sales',         '=Sheet1!$G$1:$H$10')

# Define a local/worksheet name.
workbook.define_name('Sheet2!Sales', '=Sheet2!$G$1:$G$10')

# Write some text in the file and one of the defined names in a formula.
workbook.worksheets.each do |worksheet|
  worksheet.set_column('A:A', 45)
  worksheet.write('A1', 'This worksheet contains some defined names.')
  worksheet.write('A2', 'See Formulas -> Name Manager above.')
  worksheet.write('A3', 'Example formula in cell B3 ->')

  worksheet.write('B3', '=Exchange_rate')
end

workbook.close
