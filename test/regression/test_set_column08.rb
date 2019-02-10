# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSetColumn08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_set_column08
    @xlsx = 'set_column08.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    bold        = workbook.add_format(:bold   => 1)
    italic      = workbook.add_format(:italic => 1)

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', 'Foo', bold)
    worksheet.write('B1', 'Bar', italic)
    worksheet.write('A2', data)

    worksheet.set_row(12, nil, nil, 1)
    worksheet.set_column('F:F', nil, nil, 1)

    worksheet.insert_image('E12',
                           File.join(@test_dir, 'regression', 'images/logo.png'))

    workbook.close

    compare_for_regression

  end
end
