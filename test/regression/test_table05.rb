# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionTable05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_table05
    @xlsx = 'table05.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    # Set the column width to match the taget worksheet.
    worksheet.set_column('C:F', 10.288)

    # Add the table.
    worksheet.add_table('C3:F13')

    # Add a link to check rId handling.
    worksheet.write('A1', 'http://perl.com/')

    # Add comments to check rId handling.
    worksheet.comments_author = 'John'
    worksheet.write_comment('H1', 'Test1')
    worksheet.write_comment('J1', 'Test2')

    # Add drawing to check rId handling.
    worksheet.insert_image('A4', File.join(@test_dir, 'regression', 'images/blue.png'))

    workbook.close
    compare_for_regression
  end
end
