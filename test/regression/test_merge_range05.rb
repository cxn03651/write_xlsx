# -*- coding: utf-8 -*-

require 'helper'

class TestMergeRange05 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_merge_range05
    @xlsx = 'merge_range05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(align: 'center')

    worksheet.merge_range(1, 1, 1, 3, 123, format)

    workbook.close
    compare_for_regression
  end
end
