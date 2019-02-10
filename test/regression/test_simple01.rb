# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSimple01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_simple01
    @xlsx = 'simple01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Hello')
    worksheet.write('A2', 123)

    workbook.close
    compare_for_regression
  end
end
