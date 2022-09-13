# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionComment15 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment15
    @xlsx = 'comment15.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet
    format1   = workbook.add_format(:bold => 1)

    worksheet.write('A1', 'Foo', format1)
    worksheet.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression
  end
end
