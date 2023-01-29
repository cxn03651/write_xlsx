# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionDynamicArray03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_dynamic_array03
    @xlsx = 'dynamic_array03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_formula('A1', '=1+_xlfn.XOR(1)', nil, 2)

    workbook.close
    compare_for_regression
  end
end
