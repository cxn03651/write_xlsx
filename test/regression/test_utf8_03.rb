# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_03
    @xlsx = 'utf8_03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet('Café')

    worksheet.write('A1', 'Café')

    workbook.close
    compare_for_regression
  end
end
