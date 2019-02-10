# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table11
    @xlsx = 'table11.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    data = [
            ['Foo', 1234, 2000, 4321],
            ['Bar', 1256, 4000, 4320],
            ['Baz', 2234, 3000, 4332],
            ['Bop', 1324, 1000, 4333]
           ]

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C2:F6', {:data => data})

    workbook.close
    compare_for_regression(
                                nil,
                                {  'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
