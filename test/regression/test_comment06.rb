# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionComment06 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment06
    @xlsx = 'comment06.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_comment('A1', 'Some text')
    worksheet.write_comment('A2', 'Some text')
    worksheet.write_comment('A3', 'Some text', visible: 1)
    worksheet.write_comment('A4', 'Some text')
    worksheet.write_comment('A5', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
