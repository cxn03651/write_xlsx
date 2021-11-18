# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat17 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format17
    @xlsx = 'format17.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    pattern   = workbook.add_format(:pattern => 2, :fg_color => 'red')

    worksheet.write('A1', '', pattern)

    workbook.close
    compare_for_regression
  end
end
