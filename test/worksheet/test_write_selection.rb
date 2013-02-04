# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSelection < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_selection
    @worksheet.__send__('write_selection', nil, 'A1', 'A1')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<selection activeCell="A1" sqref="A1"/>'
    assert_equal(expected, result)
  end
end
