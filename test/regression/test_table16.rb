# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable16 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table16
    @xlsx = 'table02.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    # Set the column width to match the taget worksheet.
    worksheet1.set_column('B:J', 10.288)
    worksheet2.set_column('C:L', 10.288)

    # Add the tables in reverse order to test_table02.rb
    worksheet2.add_table('I4:L11')
    worksheet2.add_table('C16:H23')

    worksheet1.add_table('B3:E11')
    worksheet1.add_table('G10:J16')
    worksheet1.add_table('C18:F25')

    workbook.close
    compare_for_regression(
                                nil,
                                {  'xl/workbook.xml' => ['<workbookView'] }
                                )
  end
end
