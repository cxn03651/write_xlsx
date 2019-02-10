# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSetColumn03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_set_column03
    @xlsx = 'set_column03.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    bold   = workbook.add_format(:bold   => 1)
    italic = workbook.add_format(:italic => 1)

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', 'Foo', italic)
    worksheet.write('B1', 'Bar', bold)
    worksheet.write('A2', data)

    worksheet.set_column('F:F', nil, bold)

    workbook.close

    compare_for_regression

  end
end
