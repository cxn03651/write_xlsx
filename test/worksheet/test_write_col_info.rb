# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteColInfo < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_col_info_1_3_5_nil_false_0_0
    min       = 1
    max       = 3
    width     = 5
    format    = nil
    hidden    = 0
    level     = 0
    collapsed = 0
    @worksheet.__send__('write_col_info', [min, max, width, format, hidden])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<col min="2" max="4" width="5.7109375" customWidth="1"/>'
    assert_equal(expected, result)
  end

  def test_write_col_info_5_5_8_nil_true_0_0
    min       = 5
    max       = 5
    width     = 8
    format    = nil
    hidden    = true
    level     = 0
    collapsed = 0
    @worksheet.__send__('write_col_info', [min, max, width, format, hidden])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>'
    assert_equal(expected, result)
  end

  def test_write_col_info_7_7_nil_1_false_0_0
    min       = 7
    max       = 7
    width     = nil
    format    = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1)
    hidden    = false
    level     = 0
    collapsed = 0
    @worksheet.__send__('write_col_info', [min, max, width, format, hidden])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<col min="8" max="8" width="9.140625" style="1"/>'
    assert_equal(expected, result)
  end

  def test_write_col_info_8_8_843_1_false_0_0
    min       = 8
    max       = 8
    width     = 8.43
    format    = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1)
    hidden    = false
    level     = 0
    collapsed = 0
    @worksheet.__send__('write_col_info', [min, max, width, format, hidden])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<col min="9" max="9" width="9.140625" style="1"/>'
    assert_equal(expected, result)
  end

  def test_write_col_info_9_9_2_nil_false_0_0
    min       = 9
    max       = 9
    width     = 2
    format    = nil
    hidden    = false
    level     = 0
    collapsed = 0
    @worksheet.__send__('write_col_info', [min, max, width, format, hidden])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<col min="10" max="10" width="2.7109375" customWidth="1"/>'
    assert_equal(expected, result)
  end

  def test_write_col_info_11_1_nil_nil_true_0_0
    min       = 11
    max       = 11
    width     = nil
    format    = nil
    hidden    = true
    level     = 0
    collapsed = 0
    @worksheet.__send__('write_col_info', [min, max, width, format, hidden])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<col min="12" max="12" width="0" hidden="1" customWidth="1"/>'
    assert_equal(expected, result)
  end
end
