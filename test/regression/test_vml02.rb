# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionVml02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_vml02
    @xlsx = 'vml02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('B2',  'Some text')
    worksheet.write_comment('D17', 'More text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    worksheet.insert_button('C4', {})
    worksheet.insert_button('E8', {})

    workbook.close
    compare_for_regression
  end
end
