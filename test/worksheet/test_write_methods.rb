# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteMethods < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write(0)
    end
  end

  def test_write_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write(0, 0, 1)
    end

    # valid cell only not raised. (but ignored)
    assert_nothing_raised do
      @worksheet.write(0, 0)
    end
    assert_nothing_raised do
      @worksheet.write('A1')
    end
    assert_nothing_raised do
      @worksheet.write(0, 0, nil)
    end
  end

  def test_write_number_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(0, 0, nil)
    end
  end

  def test_write_number_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write_number(0, 0, 1)
    end
  end

  def test_write_string_with_non_string_value
    assert_nothing_raised do
      @worksheet.write_string(0, 0, 1)
    end
  end

  def test_write_string_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(0, 0, nil)
    end
  end

  def test_write_string_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write_string(0, 0, "WriteXLSX")
    end
  end

  def test_write_rich_string_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string(0, 0, nil)
    end
  end

  def test_write_rich_string_with_valid_arg_not_raise
    format = @workbook.add_format(:bold => 1)
    assert_nothing_raised do
      @worksheet.write_rich_string(0, 0, "WriteXLSX", format, 'bold')
    end
  end

  def test_write_blank_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_blank()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_blank(0)
    end
  end

  def test_write_blank_with_valid_arg_not_raise
    format = @workbook.add_format
    assert_nothing_raised do
      @worksheet.write_blank(0, 0, format)
    end
  end

  def test_write_array_formula_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0, 1, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula('A1:B3')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0, 1, 1, nil)
    end
  end

  def test_write_array_formula_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write_array_formula(0, 0, 2, 0, '{=TREND(C1:C3,B1:B3)}')
    end
  end

  def test_write_url_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(0, 0, nil)
    end
  end

  def test_write_url_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write_url(0, 0, "http://foo.bar.com")
    end
  end

  def test_write_date_time_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time(0, 0, nil)
    end
  end

  def test_write_date_time_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write_date_time(0, 0, '2001-01-01T12:20')
    end
  end

  def test_write_comment_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment(0, 0, nil)
    end
  end

  def test_write_comment_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.write_comment(0, 0, 'comment')
    end
  end

  def test_insert_chart_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart(0, 0, nil)
    end
  end

  def test_coditional_formatting_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, nil)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting('A1:B2')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1, 1, nil)
    end
  end

  def test_conditional_formatting_with_valid_arg_not_raise
    param = {
      :type     => 'cell',
      :criteria => 'greater than',
      :value    => 5,
      :format   => $red_format
    }
    assert_nothing_raised do
      @worksheet.conditional_formatting(0, 0, param)
    end
    param = {
      :type     => 'cell',
      :criteria => 'greater than',
      :value    => 5,
      :format   => $red_format
    }
    assert_nothing_raised do
      @worksheet.conditional_formatting(0, 0, 1, 1, param)
    end
  end

  def test_data_validation_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation()
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, nil)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation('A1:B2')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1, 1)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1, 1, nil)
    end
  end

  def test_data_validation_with_valid_arg_not_raise
    param = {
      :validate => 'integer',
      :criteria => '>',
      :value    => 100
    }

    assert_nothing_raised do
      @worksheet.data_validation(0, 0, param)
    end
    assert_nothing_raised do
      @worksheet.data_validation(0, 0, 1, 1, param)
    end
  end
end
