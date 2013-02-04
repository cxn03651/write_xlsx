# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteBrk < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_brk
    @worksheet.__send__('write_brk', 1, 16383)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<brk id="1" max="16383" man="1"/>'
    assert_equal(expected, result)
  end
end
