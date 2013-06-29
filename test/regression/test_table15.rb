# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable15 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    File.delete(@xlsx) if File.exist?(@xlsx)
  end

  def test_table15
    @xlsx = 'table15.xlsx'
    workbook  = WriteXLSX.new(@xlsx)
    worksheet = workbook.add_worksheet

    data = [
            [ 'Foo', 1234, 2000, 4321 ],
            [ 'Bar', 1256,    0, 4320 ],
            [ 'Baz', 2234, 3000, 4332 ],
            [ 'Bop', 1324, 1000, 4333 ]
           ]

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C2:F6', { :data    => data })

    workbook.close
    compare_xlsx_for_regression(File.join(@regression_output, @xlsx), @xlsx,
                                nil,
                                {  'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
