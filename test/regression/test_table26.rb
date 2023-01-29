# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable26 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table26
    @xlsx = 'table26.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the target worksheet.
    worksheet.set_column('C:D', 10.288)
    worksheet.set_column('F:G', 10.288)

    # Add the table.
    worksheet.add_table('C2:D3')
    worksheet.add_table('F3:G3', header_row: 0)

    # These tables should be ignored since the ranges are incorrect.
    assert_raises(RuntimeError) do
      worksheet.add_table('I2:J2')
    end

    assert_raises(RuntimeError) do
      worksheet.add_table('L3:M3', header_row: 1)
    end

    workbook.close
    compare_for_regression
  end
end
