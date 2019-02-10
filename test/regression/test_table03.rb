# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table03
    @xlsx = 'table03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C3:F13')

    # Add a link to check rId handling.
    worksheet.write('A1', 'http://perl.com/')

    workbook.close
    compare_for_regression
  end
end
