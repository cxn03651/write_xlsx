# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_utf8_04
    @xlsx = 'utf8_04.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet('Café & Café')

    worksheet.write('A1', 'Café & Café')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
