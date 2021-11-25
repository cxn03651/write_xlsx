# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionBackground05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background05
    @xlsx = 'background05.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.set_background('test/regression/images/logo.jpg')
    worksheet2.set_background('test/regression/images/red.jpg')

    workbook.close
    compare_for_regression
  end
end
