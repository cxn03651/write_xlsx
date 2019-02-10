# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionHide01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_hide01
    @xlsx = 'hide01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet2.hide

    workbook.close
    compare_for_regression
  end
end
