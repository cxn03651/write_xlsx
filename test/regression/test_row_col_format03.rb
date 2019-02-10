# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format03
    @xlsx = 'row_col_format03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    italic    = workbook.add_format(:italic => 1)

    worksheet.set_column('A:A', nil, italic)

    workbook.close
    compare_for_regression
  end
end
