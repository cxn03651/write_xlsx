# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionIgnoreError02 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_ignore_error02
    @xlsx = 'ignore_error02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string('A1', '123')
    worksheet.ignore_errors(number_stored_as_text: 'A1')

    workbook.close
    compare_for_regression
  end
end
