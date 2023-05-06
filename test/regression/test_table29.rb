# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable29 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table29
    @xlsx = 'table29.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the target worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C3:F13')

    worksheet.insert_image(0, 0, 'test/regression/images/red.png')

    workbook.close
    compare_for_regression
  end
end
