# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteCustomFilter < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_custom_filter
    @worksheet.__send__('write_custom_filter', 4, 3000)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<customFilter operator="greaterThan" val="3000"/>'
    assert_equal(expected, result)
  end
end
