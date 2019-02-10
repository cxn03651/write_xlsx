# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_01
    @xlsx = 'utf8_01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Это фраза на русском!')

    workbook.close
    compare_for_regression
  end
end
