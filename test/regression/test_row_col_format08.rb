# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat08 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format08
    @xlsx = 'row_col_format08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)
    mixed     = workbook.add_format(:bold => 1, :italic => 1)
    italic    = workbook.add_format(:italic => 1)

    workbook.set_default_xf_indices

    worksheet.set_row(0, nil, bold)
    worksheet.set_column('A:A', nil, italic)

    worksheet.write('A1', 'Foo', mixed)
    worksheet.write('B1', 'Foo')
    worksheet.write('A2', 'Foo')
    worksheet.write('B2', 'Foo')

    workbook.close
    compare_for_regression
  end
end
