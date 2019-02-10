# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment10 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment10
    @xlsx = 'comment10.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('B2', 'Some text', :color => 51)

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
