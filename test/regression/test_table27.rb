# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable27 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table27
    @xlsx = 'table27.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Add the table.
    worksheet.add_table('C3:F13')

    # Test autofitting the columns.
    worksheet.autofit

    workbook.close
    compare_for_regression
  end
end
