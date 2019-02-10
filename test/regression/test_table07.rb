# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table07
    @xlsx = 'table07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C3:F13', {:header_row => 0})

    worksheet.write('A1', 'Foo')

    workbook.close
    compare_for_regression
  end
end
