# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat15 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_format15
    @xlsx = 'format15.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format1   = workbook.add_format(:bold => 1)
    format2   = workbook.add_format(:bold => 1, :num_format => 0)

    worksheet.write('A1', 1, format1)
    worksheet.write('A2', 2, format2)

    workbook.close
    compare_for_regression
  end
end
