# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteMethods < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write(nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write(0)
    end
  end

  def test_write_number_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(nil, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(0, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number('A1', nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_number(0, 0, nil)
    end
  end

  def test_write_string_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(nil, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(0, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string('A1', nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_string(0, 0, nil)
    end
  end

  def test_write_rich_string_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_rich_string(0, 0, nil)
    end
  end

  def test_write_blank_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_blank(nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_blank(0)
    end
  end

  def test_write_array_formula_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0, 1, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula('A1:B3')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_array_formula(0, 0, 1, 1, nil)
    end
  end

  def test_write_url_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(nil, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(0, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url('A1', nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_url(0, 0, nil)
    end
  end

  def test_write_date_time_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_date_time(0, 0, nil)
    end
  end

  def test_write_comment_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.write_comment(0, 0, nil)
    end
  end

  def test_insert_chart_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.insert_chart(0, 0, nil)
    end
  end

  def test_coditional_formatting_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting('A1:B2')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.conditional_formatting(0, 0, 1, 1, nil)
    end
  end

  def test_data_validation_with_insufficient_args_raise_InsufficientArgumentError
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation('A1')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, nil)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation('A1:B2')
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1, 1)
    end
    assert_raises(WriteXLSXInsufficientArgumentError) do
      @worksheet.data_validation(0, 0, 1, 1, nil)
    end
  end
end
