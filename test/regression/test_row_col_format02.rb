# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_row_col_format02
    @xlsx = 'row_col_format02.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)

    worksheet.set_row(0, nil, bold)
    worksheet.write('A1', 'Foo')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
