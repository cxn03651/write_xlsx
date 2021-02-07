# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteRowOff < Minitest::Test
  def setup
    @drawing = Writexlsx::Drawings.new
  end

  def test_write_row_off
    expected = '<xdr:rowOff>104775</xdr:rowOff>'

    @drawing.__send__(:write_row_off, 104775)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
