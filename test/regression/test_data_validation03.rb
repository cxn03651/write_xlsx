# -*- coding: utf-8 -*-
require 'helper'

class TestDataValidation03 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def teardown
    @tempfile.close(true)
  end

  def test_data_validation03
    @xlsx = 'data_validation03.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    worksheet.data_validation('C2',
                              validate:      'list',
                              value:         ['Foo', 'Bar', 'Baz'],
                              input_title:   'This is the input title',
                              input_message: 'This is the input message'
                              )

    values = [
              "Foobar", "Foobas", "Foobat", "Foobau", "Foobav", "Foobaw", "Foobax",
              "Foobay", "Foobaz", "Foobba", "Foobbb", "Foobbc", "Foobbd", "Foobbe",
              "Foobbf", "Foobbg", "Foobbh", "Foobbi", "Foobbj", "Foobbk", "Foobbl",
              "Foobbm", "Foobbn", "Foobbo", "Foobbp", "Foobbq", "Foobbr", "Foobbs",
              "Foobbt", "Foobbu", "Foobbv", "Foobbw", "Foobbx", "Foobby", "Foobbz",
              "Foobca", "End"
             ]

    worksheet.data_validation('D6',
                              validate:      'list',
                              value:         values,
                              input_title:   'This is the longest input title1',
                              input_message: 'This is the longest input message ' + "a"*221
                              )

    workbook.close
    compare_for_regression
  end
end
