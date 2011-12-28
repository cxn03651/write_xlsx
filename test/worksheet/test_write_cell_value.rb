# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteCellValue < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_cell_value
    @worksheet.__send__('write_cell_value', 1)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<v>1</v>'
    assert_equal(expected, result)
  end

  def test_write_cell_value_without_parameter
    @worksheet.__send__('write_cell_value')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<v></v>'
    assert_equal(expected, result)
  end

  def test_write_cell_value_with_null_string
    @worksheet.__send__('write_cell_value', '')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<v></v>'
    assert_equal(expected, result)
  end
end
