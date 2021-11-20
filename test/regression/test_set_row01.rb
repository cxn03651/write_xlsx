# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSetRow01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_set_row01
    @xlsx = 'set_row01.xlsx'
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
    worksheet.set_row(20, 15.75)
    worksheet.set_row(21, 16.50)

    workbook.close
    compare_for_regression
  end
end
