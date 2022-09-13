# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionUtf8_11 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_11
    @xlsx = 'utf8_11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', '１２３４５')

    workbook.close
    compare_for_regression
  end
end
