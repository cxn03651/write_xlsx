# -*- coding: utf-8 -*-

require 'helper'

class TestRegressionIgnoreError03 < Minitest::Test
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_ignore_error03
    @xlsx = 'ignore_error03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    10.times do |row|
      worksheet.write_string(row, 0, '123')
    end
    worksheet.ignore_errors(number_stored_as_text: 'A1:A10')

    workbook.close
    compare_for_regression
  end
end
