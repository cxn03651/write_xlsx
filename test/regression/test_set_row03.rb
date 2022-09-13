# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionSetRow03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_set_row03
    @xlsx = 'set_row03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(0, 0.75)
    worksheet.set_row(1, 1.50)
    worksheet.set_row(2, 2.25)
    worksheet.set_row(3, 3)

    worksheet.set_row(11,  9)
    worksheet.set_row(12,  9.75)
    worksheet.set_row(13, 10.50)
    worksheet.set_row(14, 11.25)

    worksheet.set_row(18, 14.25)
    worksheet.set_row(20, 15.75, nil, 1)
    worksheet.set_row(21, 16.50)

    workbook.close
    compare_for_regression
  end
end
