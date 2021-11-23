# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDynamicArray01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_dynamic_array01
    @xlsx = 'dynamic_array01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_dynamic_array_formula('A1:A1', '=AVERAGE(TIMEVALUE(B1:B2))', nil, 0)
    worksheet.write('B1', '12:00')
    worksheet.write('B2', '12:00')

    workbook.close
    compare_for_regression
  end
end
