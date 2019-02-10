# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format07
    @xlsx = 'row_col_format07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    worksheet.set_row(4, nil, bold)

    workbook.close
    compare_for_regression
  end
end
