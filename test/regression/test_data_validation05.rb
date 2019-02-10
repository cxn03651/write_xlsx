# -*- coding: utf-8 -*-
require 'helper'

class TestDataValidation05 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_data_validation05
    @xlsx = 'data_validation02.xlsx'
    workbook  = WriteXLSX.new(@io)
    worksheet = workbook.add_worksheet

    values = [
              "Foobar", "Foobas", "Foobat", "Foobau", "Foobav", "Foobaw", "Foobax",
              "Foobay", "Foobaz", "Foobba", "Foobbb", "Foobbc", "Foobbd", "Foobbe",
              "Foobbf", "Foobbg", "Foobbh", "Foobbi", "Foobbj", "Foobbk", "Foobbl",
              "Foobbm", "Foobbn", "Foobbo", "Foobbp", "Foobbq", "Foobbr", "Foobbs",
              "Foobbt", "Foobbu", "Foobbv", "Foobbw", "Foobbx", "Foobby", "Foobbz",
              "Foobca", "End"
             ]
    input_message = 'a' * 256
    e = assert_raise(RuntimeError) do
      worksheet.data_validation('D6',
                                validate:      'list',
                                value:         values,
                                input_title:   'This is the longest input title',
                                input_message: input_message.dup
                                )
    end
    message = e.message
    assert_equal("Length of input message '#{input_message}' exceeds Excel's limit of 255",
                 message)
    workbook.close
  end
end
