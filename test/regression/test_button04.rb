# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_button04
    @xlsx = 'button04.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_button('C2', {})
    worksheet2.insert_button('E5', {})

    workbook.close
    compare_for_regression
  end
end
