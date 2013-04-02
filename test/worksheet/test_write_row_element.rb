# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteRow < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_row_0
    @worksheet.__send__('write_row_element', 0){}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="1"></row>'
    assert_equal(expected, result)
  end

  def test_write_row_2_22
    @worksheet.__send__('write_row_element', 2, '2:2'){}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="3" spans="2:2"></row>'
    assert_equal(expected, result)
  end

  def test_write_row_1_nil_30
    @worksheet.__send__('write_row_element', 1, nil, 30){}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="2" ht="30" customHeight="1"></row>'
    assert_equal(expected, result)
  end

  def test_write_row_3_nil_nil_nil_1
    @worksheet.__send__('write_row_element', 3, nil, nil, nil, 1){}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="4" hidden="1"></row>'
    assert_equal(expected, result)
  end

  def test_write_row_6_nil_nil_format
    format = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1)
    @worksheet.__send__('write_row_element', 6, nil, nil, format){}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="7" s="1" customFormat="1"></row>'
    assert_equal(expected, result)
  end

  def test_write_row_9_nil_3
    @worksheet.__send__('write_row_element', 9, nil, 3) {}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="10" ht="3" customHeight="1"></row>'
    assert_equal(expected, result)
  end

  def test_write_row_12_nil_24_nil_1
    @worksheet.__send__('write_row_element', 12, nil, 24, nil, 1){}
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="13" ht="24" hidden="1" customHeight="1"></row>'
    assert_equal(expected, result)
  end

  def test_write_empty_row_12_nil_24_nil_1
    @worksheet.__send__('write_empty_row', 12, nil, 24, nil, 1)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<row r="13" ht="24" hidden="1" customHeight="1"/>'
    assert_equal(expected, result)
  end
end
