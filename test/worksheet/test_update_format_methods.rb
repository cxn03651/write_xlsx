# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestUpdateFormatMethods < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_update_format_with_params_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_format_with_params
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_format_with_params(0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_format_with_params('A1')
    end
  end

  def test_update_format_with_params_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.update_format_with_params(0, 0, color: 'red', border: 2)
    end
    assert_nothing_raised do
      @worksheet.update_format_with_params('B2', align: 'center')
    end
  end

  def test_update_format_with_params_should_write_blank_when_there_is_no_CellData
    assert_nil(@worksheet.instance_variable_get(:@cell_data_table)[0])
    @worksheet.update_format_with_params(0, 0, left: 4)
    assert @worksheet.instance_variable_get(:@cell_data_table)[0][0] != nil
  end

  def test_update_format_with_params_should_keep_data_when_updating_format
    number = 153
    @worksheet.write(0, 0, number)
    @worksheet.update_format_with_params(0, 0, bg_color: 'gray')
    assert_equal(@worksheet.instance_variable_get(:@cell_data_table)[0][0].data, number)

    string = 'Hello, World!'
    @worksheet.write(0, 0, string)
    @worksheet.update_format_with_params(0, 0, bg_color: 'gray')
    written_string = @workbook.shared_strings.string(@worksheet.instance_variable_get(:@cell_data_table)[0][0].data[:sst_id])
    assert_equal(written_string, string)

    formula = '=1+1'
    @worksheet.write(0, 0, formula)
    @worksheet.update_format_with_params(0, 0, bg_color: 'gray')
    assert_equal(@worksheet.instance_variable_get(:@cell_data_table)[0][0].token, '1+1')

    array_formula = '{=SUM(B1:C1*B2:C2)}'
    @worksheet.write('A1', array_formula)
    @worksheet.update_format_with_params(0, 0, bg_color: 'gray')
    assert_equal(@worksheet.instance_variable_get(:@cell_data_table)[0][0].token, 'SUM(B1:C1*B2:C2)')

    url = 'https://www.writexlsx.io'
    @worksheet.write(0, 0, url)
    @worksheet.update_format_with_params(0, 0, bg_color: 'gray')
    written_string = @workbook.shared_strings.string(@worksheet.instance_variable_get(:@cell_data_table)[0][0].data[:sst_id])
    assert_equal(written_string, url)

    string = 'Hello, World!'
    format = @workbook.add_format(color: 'white')
    @worksheet.write(0, 0, string, format)
    @worksheet.update_format_with_params(0, 0, bg_color: 'gray')
    written_string = @workbook.shared_strings.string(@worksheet.instance_variable_get(:@cell_data_table)[0][0].data[:sst_id])
    assert_equal(written_string, string)
  end

  def test_update_range_format_with_params_with_insufficient_args_raise_InsufficientArgumentError
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_range_format_with_params
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_range_format_with_params(0, 0)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_range_format_with_params('A1')
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_range_format_with_params(0, 0, 3, 3)
    end
    assert_raise(WriteXLSXInsufficientArgumentError) do
      @worksheet.update_range_format_with_params('A2:C2')
    end
  end

  def test_update_range_format_with_params_with_valid_arg_not_raise
    assert_nothing_raised do
      @worksheet.update_range_format_with_params(0, 0, 3, 3, bold: 1)
    end
    assert_nothing_raised do
      @worksheet.update_range_format_with_params('A2:D5', bold: 1)
    end
  end
end
