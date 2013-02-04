# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteFilterColumn < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_filter_column
    @worksheet.__send__('write_filter_column', 0, 1, *['East'])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<filterColumn colId="0"><filters><filter val="East"/></filters></filterColumn>'
    assert_equal(expected, result)
  end
end
