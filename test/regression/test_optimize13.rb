# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionOptimize13 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_optimize13
    @xlsx = 'optimize13.xlsx'
    workbook  = WriteXLSX.new(@io)
    workbook.set_optimization
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
