# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment03
    @xlsx = 'comment03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('A1',         'Some text')
    worksheet.write_comment('XFD1048576', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
