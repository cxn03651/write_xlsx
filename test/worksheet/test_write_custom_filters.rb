# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteCustomFilters < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_custom_filters_4_4000
    @worksheet.__send__('write_custom_filters', 4, 4000)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<customFilters><customFilter operator="greaterThan" val="4000"/></customFilters>'
    assert_equal(expected, result)
  end

  def test_write_custom_filters_4_3000_0_1_8000
    @worksheet.__send__('write_custom_filters', 4, 3000, 0, 1, 8000)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<customFilters and="1"><customFilter operator="greaterThan" val="3000"/><customFilter operator="lessThan" val="8000"/></customFilters>'
    assert_equal(expected, result)
  end
end
