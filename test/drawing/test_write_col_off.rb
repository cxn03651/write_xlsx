# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteColOff < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_col_off
    expected = '<xdr:colOff>457200</xdr:colOff>'

    @drawing.__send__(:write_col_off, 457200)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
