# -*- coding: utf-8 -*-

require 'helper'

class TestDataValidation08 < Minitest::Test
  def setup
    setup_dir_var
  end

  def test_data_validation08
    @xlsx = 'data_validation08.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.data_validation(
      'C2',
      validate: 'any',
      input_title: 'This is the input title',
      input_message: 'This is the input message'
    )
    workbook.close

    compare_for_regression
  end
end
