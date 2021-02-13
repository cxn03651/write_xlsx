# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionComment16 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_comment16
    @xlsx = 'comment16.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write('A1',  'Foo')
    worksheet.write('C7',  'Bar')
    worksheet.write('G14', 'Baz')

    worksheet.write_comment('A1',  'Some text')
    worksheet.write_comment('D1',  'Some text')
    worksheet.write_comment('C7',  'Some text')
    worksheet.write_comment('E10', 'Some text')
    worksheet.write_comment('G14', 'Some text')

    # Set the author to match the target XLSX file.
    worksheet.comments_author = 'John'

    workbook.close
    compare_for_regression
  end
end
