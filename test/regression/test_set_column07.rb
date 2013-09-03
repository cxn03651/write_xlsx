# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSetColumn07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_set_column07
    @xlsx = 'set_column07.xlsx'
    workbook    = WriteXLSX.new(@xlsx)
    worksheet   = workbook.add_worksheet

    bold        = workbook.add_format(:bold   => 1)
    italic      = workbook.add_format(:italic => 1)
    bold_italic = workbook.add_format(:bold   => 1, :italic => 1)

    data = [
            [1, 2, 3,  4,  5],
            [2, 4, 6,  8, 10],
            [3, 6, 9, 12, 15]
           ]

    worksheet.write('A1', 'Foo', italic)
    worksheet.write('B1', 'Bar', bold)
    worksheet.write('A2', data)

    worksheet.set_row(12, nil, italic)
    worksheet.set_column('F:F', nil, bold)

    worksheet.write('F13', nil, bold_italic)

    worksheet.insert_image('E12',
                           File.join(@test_dir, 'regression', 'images/logo.png'))

    workbook.close

    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)

  end
end
