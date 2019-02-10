# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionRowColFormat09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_row_col_format09
    @xlsx = 'row_col_format09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    bold      = workbook.add_format(:bold => 1)
    mixed     = workbook.add_format(:bold => 1, :italic => 1)
    italic    = workbook.add_format(:italic => 1)

    workbook.set_default_xf_indices

    worksheet.set_row(4, nil, bold)
    worksheet.set_column('C:C', nil, italic)

    worksheet.write('C1', 'Foo')
    worksheet.write('A5', 'Foo')
    worksheet.write('C5', 'Foo', mixed)

    workbook.close
    compare_for_regression
  end
end
