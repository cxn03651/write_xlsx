# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionDynamicArray04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_dynamic_array04
    @xlsx = 'dynamic_array04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    worksheet.write_dynamic_array_formula(
      'A1', '=AVERAGE(TIMEVALUE(B1:B2))', bold, 0.5
    )
    worksheet.write('B1', '12:00')
    worksheet.write('B2', '12:00')

    workbook.close
    compare_for_regression
  end
end
