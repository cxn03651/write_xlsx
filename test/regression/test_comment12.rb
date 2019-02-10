# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment12 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment12
    @xlsx = 'comment12.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_row(0, 21)
    worksheet.set_column('B:B', 10)

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('A1', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression
  end
end
