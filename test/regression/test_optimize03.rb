# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionOptimize03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_optimize03
    @xlsx = 'optimize03.xlsx'
    workbook  = WriteXLSX.new(@io)
    workbook.set_optimization
    worksheet = workbook.add_worksheet

    bold = workbook.add_format(bold: 1)

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

    hyperlink_format = workbook.add_format(
      color:     'blue',
      underline: 1
    )

    headings = ['Features of Excel::Writer::XLSX', '']
    worksheet.write_row('A1', headings, heading)

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
    worksheet.write('B4', "\u0410 \u0411 \u0412 \u0413 \u0414")

    num1_format = workbook.add_format(num_format: '$#,##0.00')
    num2_format = workbook.add_format(num_format: ' d mmmm yyy')

    worksheet.write('A5', "Numbers")
    worksheet.write('B5', 1234.56)
    worksheet.write('A6', "Formatted numbers")
    worksheet.write('B6', 1234.56, num1_format)
    worksheet.write('A7', "Formatted numbers")
    worksheet.write('B7', 37257, num2_format)

    worksheet.set_selection('B8')
    worksheet.write('A8', 'Formulas and functions, "=SIN(PI()/4)"')
    worksheet.write('B8', '=SIN(PI()/4)')

    worksheet.write('A9', "Hyperlinks")
    worksheet.write('B9', 'http://www.perl.com/', hyperlink_format)

    workbook.close
    compare_for_regression
  end
end
