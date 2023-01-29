#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'write_xlsx'

workbook   = WriteXLSX.new('demo.xlsx')
worksheet  = workbook.add_worksheet('Demo')
worksheet2 = workbook.add_worksheet('Another sheet')
worksheet3 = workbook.add_worksheet('And another')

bold = workbook.add_format(bold: 1)

#######################################################################
#
# Write a general heading
#
worksheet.set_column('A:A', 36, bold)
worksheet.set_column('B:B', 20)
worksheet.set_row(0, 40)

heading = workbook.add_format(
  bold:  1,
  color: 'blue',
  size:  16,
  merge: 1,
  align: 'vcenter'
)

headings = ['Features of WriteXLSX', '']
worksheet.write_row('A1', headings, heading)

#######################################################################
#
# Some text examples
#
text_format = workbook.add_format(
  bold:   1,
  italic: 1,
  color:  'red',
  size:   18,
  font:   'Lucida Calligraphy'
)

worksheet.write('A2', "Text")
worksheet.write('B2', "Hello Excel")
worksheet.write('A3', "Formatted text")
worksheet.write('B3', "Hello Excel", text_format)
worksheet.write('A4', "Unicode text")
worksheet.write('B4', "А Б В Г Д")

#######################################################################
#
# Some numeric examples
#
num1_format = workbook.add_format(num_format: '$#,##0.00')
num2_format = workbook.add_format(num_format: ' d mmmm yyy')

worksheet.write('A5', "Numbers")
worksheet.write('B5', 1234.56)
worksheet.write('A6', "Formatted numbers")
worksheet.write('B6', 1234.56, num1_format)
worksheet.write('A7', "Formatted numbers")
worksheet.write('B7', 37257, num2_format)

#######################################################################
#
# Formulae
#
worksheet.set_selection('B8')
worksheet.write('A8', 'Formulas and functions, "=SIN(PI()/4)"')
worksheet.write('B8', '=SIN(PI()/4)')

#######################################################################
#
# Hyperlinks
#
worksheet.write('A9', "Hyperlinks")
worksheet.write('B9', 'http://www.ruby-lang.org/')

#######################################################################
#
# Images
#
worksheet.write('A10', "Images")
worksheet.insert_image(
  'B10', 'republic.png',
  x_offset: 16, y_offset: 8
)

#######################################################################
#
# Misc
#
worksheet.write('A18', "Page/printer setup")
worksheet.write('A19', "Multiple worksheets")

workbook.close
