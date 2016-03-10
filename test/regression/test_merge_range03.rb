# -*- coding: utf-8 -*-
require 'helper'

class TestMergeRange03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_merge_range03
    @xlsx = 'merge_range03.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    worksheet.merge_range(1, 1, 1, 2, 'Foo', format)
    worksheet.merge_range(1, 3, 1, 4, 'Foo', format)
    worksheet.merge_range(1, 5, 1, 6, 'Foo', format)

    workbook.close
    compare_xlsx_for_regression(
      File.join(@regression_output, @xlsx),
      @xlsx
    )
  end
end
