# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionOptimize01 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_optimize01
    @xlsx = 'optimize01.xlsx'
    workbook  = WriteXLSX.new(@io)
    workbook.set_optimization
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Hello')
    worksheet.write('A2', 123)

    workbook.close
    compare_for_regression
  end

  def test_optimize01_constructor_option
    @xlsx = 'optimize01.xlsx'
    workbook  = WriteXLSX.new(@io, optimization: true)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Hello')
    worksheet.write('A2', 123)

    workbook.close
    compare_for_regression
  end
end
