# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable30 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table30
    @xlsx = 'table30.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the target worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C3:F13')

    worksheet.set_background('test/regression/images/logo.jpg')

    workbook.close
    compare_for_regression
  end
end
