# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_utf8_09
    @xlsx = 'utf8_09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    num_format = workbook.add_format(:num_format => '[$¥-411]#,##0.00')

    worksheet.write('A1', 1, num_format)

    workbook.close
    compare_for_regression
  end
end
