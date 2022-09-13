# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/drawing'

class TestWriteXfrmOffset < Minitest::Test
  def setup
    @drawing = Writexlsx::Drawings.new
  end

  def test_write_xfrm_offset
    expected = '<a:off x="0" y="0"/>'

    @drawing.__send__(:write_xfrm_offset)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
