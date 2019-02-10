# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat14 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format14
    @xlsx = 'row_col_format14.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format    = workbook.add_format(:bold => 1)

    worksheet.set_column('B:D', 5)
    worksheet.set_column('F:F', 8, nil)
    worksheet.set_column('H:H', nil, format)
    worksheet.set_column('J:J', 2)
    worksheet.set_column('L:L', nil, nil, 1)

    workbook.close
    compare_for_regression
  end
end
