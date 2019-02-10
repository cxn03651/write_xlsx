# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment09 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment09
    @xlsx = 'comment09.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_comment('A1', 'Some text', :author => 'John')
    worksheet.write_comment('A2', 'Some text', :author => 'Perl')
    worksheet.write_comment('A3', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
