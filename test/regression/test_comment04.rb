# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment04
    @xlsx = 'comment04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet
    worksheet3 = workbook.add_worksheet

    worksheet1.write('A1', 'Foo')
    worksheet1.write_comment('B2', 'Some text')

    worksheet3.write('A1', 'Bar')
    worksheet3.write_comment('C7', 'More text')

    # Set the author to match the target XLSX file.
    worksheet1.comments_author = 'John'
    worksheet3.comments_author = 'John'

    workbook.close
    compare_for_regression(
      nil,
      { 'xl/workbook.xml' => ['<workbookView'] }
    )
  end
end
