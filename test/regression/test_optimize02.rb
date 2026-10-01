# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionOptimize02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_optimize02
    @xlsx = 'optimize02.xlsx'
    workbook  = WriteXLSX.new(@io)
    workbook.set_optimization
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Hello')
    worksheet.write('A2', 123)

    # This should be ignored since later row has already been written.
    worksheet.write('G1', 'Foo')

    workbook.close
    compare_for_regression
  end
end
