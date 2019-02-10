# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionUtf8_07 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_utf8_07
    @xlsx = 'utf8_07.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1', 'Foo')
    worksheet.write_comment('A1', 'Café')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression
  end
end
