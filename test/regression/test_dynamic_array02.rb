# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionDynamicArray02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_dynamic_array02
    @xlsx = 'dynamic_array02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('B1', '=UNIQUE(A1)', nil, 0)
    worksheet.write('A1', 0)

    workbook.close
    compare_for_regression
  end
end
