# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionFormat12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_format12
    @xlsx = 'format12.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    top_left_bottom = workbook.add_format(
      :left   => 1,
      :top    => 1,
      :bottom => 1
    )

    top_bottom = workbook.add_format(
      :top    => 1,
      :bottom => 1
    )

    top_left = workbook.add_format(
      :left   => 1,
      :top    => 1
    )

    worksheet.write('B2', 'test', top_left_bottom)
    worksheet.write('D2', 'test', top_left)
    worksheet.write('F2', 'test', top_bottom)

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
