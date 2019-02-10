# -*- coding: utf-8 -*-
require 'helper'

class TestDataValidation04 < Test::Unit::TestCase
  def setup
    setup_dir_var
  end

  def test_data_validation04
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

    input_title = 'a' * 33
    e = assert_raise(RuntimeError) do
      worksheet.data_validation('D6',
                                validate:      'list',
                                value:         values,
                                input_title:   input_title.dup,
                                input_message: 'This is the longest input message ' + "a"*221
                                )
    end
    message = e.message
    assert_equal("Length of input title '#{input_title}' exceeds Excel's limit of 32",
                 message)
    workbook.close
  end
end
