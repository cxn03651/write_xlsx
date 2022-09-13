# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionSetRow02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_set_row02
    @xlsx = 'set_row01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row_pixels(0, 1)
    worksheet.set_row_pixels(1, 2)
    worksheet.set_row_pixels(2, 3)
    worksheet.set_row_pixels(3, 4)

    worksheet.set_row_pixels(11, 12)
    worksheet.set_row_pixels(12, 13)
    worksheet.set_row_pixels(13, 14)
    worksheet.set_row_pixels(14, 15)

    worksheet.set_row_pixels(18, 19)
    worksheet.set_row_pixels(20, 21)
    worksheet.set_row_pixels(21, 22)

    workbook.close
    compare_for_regression
  end
end
