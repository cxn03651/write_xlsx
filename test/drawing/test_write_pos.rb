# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWritePos < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_pos
    expected = '<xdr:pos x="0" y="0"/>'

    @drawing.__send__(:write_pos, 0, 0)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
