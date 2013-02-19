# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_button04
    @xlsx = 'button04.xlsx'
    workbook   = WriteXLSX.new(@xlsx)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet1.insert_button('C2', {})
    worksheet2.insert_button('E5', {})

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
