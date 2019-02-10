# -*- coding: utf-8 -*-
require 'helper'

class TestDataValidation02 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_data_validation02
    @xlsx = 'data_validation02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.data_validation('C2',
                              validate:      'list',
                              value:         ['Foo', 'Bar', 'Baz'],
                              input_title:   'This is the input title',
                              input_message: 'This is the input message'
                              )
    workbook.close
    compare_for_regression
  end
end
