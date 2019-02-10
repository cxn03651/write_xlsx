# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment11 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment11
    @xlsx = 'comment11.xlsx'
    workbook   = WriteXLSX.new(@io)
    worksheet1 = workbook.add_worksheet
    worksheet2 = workbook.add_worksheet

    worksheet2.write('A1', 'Foo')
    worksheet2.write_comment('B2', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet2.comments_author = 'John'

    workbook.close
    compare_for_regression(
      [],
      {'xl/workbook.xml' => ['<workbookView']}
    )
  end
end
