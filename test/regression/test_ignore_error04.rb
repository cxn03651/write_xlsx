# -*- coding: utf-8 -*-
require 'helper'

class TestRegressionIgnoreError04 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_ignore_error04
    @xlsx = 'ignore_error04.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.write_string('A1', '123')
    worksheet.write_string('C3', '123')
    worksheet.write_string('E5', '123')
    worksheet.ignore_errors(number_stored_as_text: 'A1 C3 E5')

    workbook.close
    compare_for_regression
  end
end
