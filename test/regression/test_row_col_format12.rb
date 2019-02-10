# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format12
    @xlsx = 'row_col_format12.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_column('C:C', nil, nil, 1)

    workbook.close
    compare_for_regression
  end
end
