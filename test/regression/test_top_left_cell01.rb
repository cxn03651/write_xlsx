# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTopLeftCell01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_top_left_cell01
    @xlsx = 'top_left_cell01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_top_left_cell(15, 0)

    workbook.close
    compare_for_regression
  end
end
