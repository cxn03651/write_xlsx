# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionBackground04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_background04
    @xlsx = 'background04.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.set_background('test/regression/images/logo.jpg')
    worksheet2.set_background('test/regression/images/logo.jpg')

    workbook.close
    compare_for_regression
  end
end
