# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionSetColumn09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_set_column09
    @xlsx = 'set_column09.xlsx'
    workbook    = WriteXLSX.new(@io)
    worksheet   = workbook.add_worksheet

    # Test the order and overwriting of columns.
    worksheet.set_column('A:A', 100)
    worksheet.set_column('F:H',   8)
    worksheet.set_column('C:D',  12)
    worksheet.set_column('A:A',  10)
    worksheet.set_column('XFD:XFD', 5)
    worksheet.set_column('ZZ:ZZ', 3)

    workbook.close

    compare_for_regression

  end
end
