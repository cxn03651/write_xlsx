# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionButton03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_button03
    @xlsx = 'button03.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    worksheet.insert_button('C2', {})
    worksheet.insert_button('E5', {})

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
