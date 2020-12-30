# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable22 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true) if @tempfile
  end

  def test_table22
    @xlsx = 'table22.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    data = [
      ['apple', 'pie' ],
      ['pine',  'tree']
    ]

    # Set the column width to match the taget worksheet.
    worksheet.set_column('B:C', 10.288)

    # Add the table.
    worksheet.add_table('B2:C3', :data => data, :header_row => 0)

    workbook.close
    compare_for_regression
  end
end
