# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteRowBreaks < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_row_breaks_1
    @worksheet.instance_variable_get(:@page_setup).hbreaks = [1]
    @worksheet.__send__('write_row_breaks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<rowBreaks count="1" manualBreakCount="1"><brk id="1" max="16383" man="1"/></rowBreaks>'
    assert_equal(expected, result)
  end

  def test_write_row_breaks_15_7_3_0
    @worksheet.instance_variable_get(:@page_setup).hbreaks = [15, 7, 3, 0]
    @worksheet.__send__('write_row_breaks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<rowBreaks count="3" manualBreakCount="3"><brk id="3" max="16383" man="1"/><brk id="7" max="16383" man="1"/><brk id="15" max="16383" man="1"/></rowBreaks>'
    assert_equal(expected, result)
  end
end
