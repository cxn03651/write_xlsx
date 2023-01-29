# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTopLeftCell03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_top_left_cell03
    @xlsx = 'top_left_cell03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_top_left_cell('AA32')

    workbook.close
    compare_for_regression
  end
end
