# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteColBreaks < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_col_breaks_1
    @worksheet.instance_variable_get(:@page_setup).vbreaks = [1]
    @worksheet.__send__('write_col_breaks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<colBreaks count="1" manualBreakCount="1"><brk id="1" max="1048575" man="1"/></colBreaks>'
    assert_equal(expected, result)
  end

  def test_write_col_breaks_8_3_1_0
    @worksheet.instance_variable_get(:@page_setup).vbreaks = [8, 3, 1, 0]
    @worksheet.__send__('write_col_breaks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<colBreaks count="3" manualBreakCount="3"><brk id="1" max="1048575" man="1"/><brk id="3" max="1048575" man="1"/><brk id="8" max="1048575" man="1"/></colBreaks>'
    assert_equal(expected, result)
  end
end
