# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteCol < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_col
    expected = '<xdr:col>4</xdr:col>'

    @drawing.__send__(:write_col, 4)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
