# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/drawing'

class TestWriteAGraphicFrameLocks < Minitest::Test
  def setup
    @drawing = Writexlsx::Drawings.new
  end

  def test_write_a_graphic_frame_locks
    expected = '<a:graphicFrameLocks noGrp="1"/>'

    @drawing.__send__(:write_a_graphic_frame_locks)
    result = @drawing.instance_variable_get(:@writer).string

    assert_equal(expected, result)
  end
end
