# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format11
    @xlsx = 'row_col_format11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('C:C', 4)

    workbook.close
    compare_for_regression
  end
end
