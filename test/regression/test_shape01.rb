# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionShape01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_shape01
    @xlsx = 'shape01.xlsx'
    workbook   = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    rect = workbook.add_shape

    worksheet.insert_shape('C2', rect)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
