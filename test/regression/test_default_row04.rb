# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionDefaultRow04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_default_row04
    @xlsx = 'default_row04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.set_default_row(24)

    worksheet.write('A1',  'Foo')
    worksheet.write('A10', 'Bar')

    worksheet.write_comment('C4', 'Hello', :y_offset => 22)

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression
  end
end
