# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_04
    @xlsx = 'utf8_04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet('Café & Café')

    worksheet.write('A1', 'Café & Café')

    workbook.close
    compare_for_regression
  end
end
