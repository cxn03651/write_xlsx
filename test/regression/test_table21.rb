# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionTable21 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table21
    @xlsx = 'table21.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Column')

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:D', 10.288)

    # Add the table.
    worksheet.add_table(
      'C3:D13',
      {
        :columns => [
          { :header => "Column" }
        ]
      }
    )

    workbook.close
    compare_for_regression
  end
end
