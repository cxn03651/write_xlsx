# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteExt < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_ext
    expected = '<xdr:ext cx="9308969" cy="6078325"/>'

    @drawing.__send__(:write_ext, 9308969, 6078325)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
