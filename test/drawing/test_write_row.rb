# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestDrawingWriteRow < Minitest::Test
  def setup
    @drawing = Writexlsx::Drawings.new
  end

  def test_write_row
    expected = '<xdr:row>8</xdr:row>'

    @drawing.__send__(:write_row, 8)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
