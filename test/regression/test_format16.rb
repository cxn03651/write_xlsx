# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat16 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format16
    @xlsx = 'format16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    pattern   = workbook.add_format(:pattern => 2)

    worksheet.write('A1', '', pattern)

    workbook.close
    compare_for_regression
  end
end
