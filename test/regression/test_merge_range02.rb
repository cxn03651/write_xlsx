# -*- coding: utf-8 -*-
require 'helper'

class TestMergeRange02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_merge_range02
    @xlsx = 'merge_range02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:align => 'center')

    worksheet.merge_range(1, 1, 5, 3, 'Foo', format)

    workbook.close
    compare_for_regression
  end
end
