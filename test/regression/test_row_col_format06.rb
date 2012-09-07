# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat06 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_row_col_format06
    @xlsx = 'row_col_format06.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)
    italic    = workbook.add_format(:italic => 1)

    worksheet.set_column('A:A', nil, bold)
    worksheet.set_column('C:C', nil, italic)

    worksheet.write('A1', 'Foo')
    worksheet.write('C1', 'Bar')

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx)
  end
end
