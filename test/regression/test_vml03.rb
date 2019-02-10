# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionVml03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_vml03
    @xlsx = 'vml03.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.write('A1', 'Foo')
    worksheet1.write_comment('B2',  'Some text')

    worksheet3.write('A1', 'Bar')
    worksheet3.write_comment('C7', 'More text')

    # Set the author to match the target XLSX file.
    worksheet1.comments_author = 'John'
    worksheet3.comments_author = 'John'

    worksheet1.insert_button('C4', {})
    worksheet1.insert_button('E8', {})

    worksheet3.insert_button('B2', {})
    worksheet3.insert_button('C4', {})
    worksheet3.insert_button('E8', {})

    workbook.close
    compare_for_regression
  end
end
