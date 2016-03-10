# -*- coding: utf-8 -*-
require 'helper'

class TestMergeRange05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_merge_range05
    @xlsx = 'merge_range05.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    worksheet.merge_range(1, 1, 1, 3, 123, format)

    workbook.close
    compare_xlsx_for_regression(
      File.join(@regression_output, @xlsx),
      @xlsx
    )
  end
end
