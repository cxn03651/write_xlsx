# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteCNvPr < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_c_nv_pr
    expected = '<xdr:cNvPr id="2" name="Chart 1"/>'

    @drawing.__send__(:write_c_nv_pr, 2, 'Chart 1')
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
