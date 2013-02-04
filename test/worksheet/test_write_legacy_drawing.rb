# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteLegacyDrawing < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_legacy_drawing
    @worksheet.write_comment('A1', 'comment')
    @worksheet.__send__('write_legacy_drawing')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<legacyDrawing r:id="rId1"/>'
    assert_equal(expected, result)
  end
end
