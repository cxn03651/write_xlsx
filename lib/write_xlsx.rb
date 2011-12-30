# -*- coding: utf-8 -*-

require 'write_xlsx/workbook'

#
# write_xlsx is gem to create a new file in the Excel 2007+ XLSX format,
# and you can use the same interface as writeexcel gem.
# write_xlsx is converted from Perl’s module github.com/jmcnamara/excel-writer-xlsx .
#
# == Description
# The WriteXLSX supports the following features:
#
#    Multiple worksheets
#    Strings and numbers
#    Unicode text
#    Rich string formats
#    Formulas (including array formats)
#    cell formatting
#    Embedded images
#    Charts
#    Autofilters
#    Data validation
#    Hyperlinks
#    Defined names
#    Grouping/Outlines
#    Cell comments
#    Panes
#    Page set-up and printing options
# WriteXLSX uses the same interface as WriteExcel gem.
#
# == Synopsis
# To write a string, a formatted string, a number and a formula to the
# first worksheet in an Excel XMLX spreadsheet called ruby.xlsx:
#
#   require 'rubygems'
#   require 'write_xlsx'
#
#   # Create a new Excel workbook
#   workbook = WriteXLSX.new('ruby.xlsx')
#
#   # Add a worksheet
#   worksheet = workbook.add_worksheet
#
#   #  Add and define a format
#   format = workbook.add_format # Add a format
#   format.set_bold
#   format.set_color('red')
#   format.set_align('center')
#
#   # Write a formatted and unformatted string, row and column notation.
#   col = row = 0
#   worksheet.write(row, col, "Hi Excel!", format)
#   worksheet.write(1,   col, "Hi Excel!")
#
#   # Write a number and a formula using A1 notation
#   worksheet.write('A3', 1.2345)
#   worksheet.write('A4', '=SIN(PI()/4)')
#   workbook.close
#
# == Other Methods
#
# see Writexlsx::Workbook, Writexlsx::Worksheet, Writexlsx::Chart etc.
#
class WriteXLSX < Writexlsx::Workbook
  if RUBY_VERSION < '1.9'
    $KCODE = 'u'
  end
end
