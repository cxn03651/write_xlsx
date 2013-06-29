# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_utf8_09
    @xlsx = 'utf8_09.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    num_format = workbook.add_format(:num_format => '[$¥-411]#,##0.00')

    worksheet.write('A1', 1, num_format)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
