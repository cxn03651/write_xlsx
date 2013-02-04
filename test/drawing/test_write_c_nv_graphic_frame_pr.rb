# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class TestWriteCNvGraphicFramePr < Test::Unit::TestCase
  def setup
    @drawing = Writexlsx::Drawing.new
  end

  def test_write_c_nv_graphic_frame_pr_01
    expected = '<xdr:cNvGraphicFramePr/>'

    @drawing.instance_variable_set(:@embedded, 1)
    @drawing.__send__(:write_c_nv_graphic_frame_pr)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end

  def test_write_c_nv_graphic_frame_pr_02
    expected = '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>'

    @drawing.__send__(:write_c_nv_graphic_frame_pr)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
