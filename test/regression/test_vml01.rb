# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionVml01 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_vml01
    @xlsx = 'vml01.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    worksheet.insert_button('C4', {})

    workbook.close
    compare_for_regression
  end
end
