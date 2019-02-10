# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format05
    @xlsx = 'row_col_format05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)
    italic    = workbook.add_format(:italic => 1)

    worksheet.set_column('A:A', nil, bold)
    worksheet.set_column('C:C', nil, italic)

    workbook.close
    compare_for_regression
  end
end
