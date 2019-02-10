# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format10
    @xlsx = 'row_col_format10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    worksheet.set_column('C:C', nil, bold)

    workbook.close
    compare_for_regression
  end
end
