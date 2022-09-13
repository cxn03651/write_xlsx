# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionBackground02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background02
    @xlsx = 'background02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_background('test/regression/images/logo.jpg')

    workbook.close
    compare_for_regression
  end
end
