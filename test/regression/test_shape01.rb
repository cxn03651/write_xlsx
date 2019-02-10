# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShape01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_shape01
    @xlsx = 'shape01.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    rect = workbook.add_shape

    worksheet.insert_shape('C2', rect)

    workbook.close
    compare_for_regression
  end
end
