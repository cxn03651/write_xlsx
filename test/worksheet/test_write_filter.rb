# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteFilter < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_filter
    @worksheet.__send__('write_filter', 'East')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<filter val="East"/>'
    assert_equal(expected, result)
  end
end
