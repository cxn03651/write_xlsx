# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'write_xlsx/shape'
require 'write_xlsx/drawing'
require 'stringio'

class DrawingShape05 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_drawing_shape_05
    shape = Writexlsx::Shape.new
    shape.id       = 1000
    shape.flip_v   = 1
    shape.flip_h   = 1
    shape.rotation = 90

    @obj = Writexlsx::Drawing.new
    @obj.instance_variable_set(:@palette, @worksheet.instance_variable_get(:@palette))
    @obj.embedded = 1

    @obj.add_drawing_object(
                            3, 4, 8, 209550, 95250, 12, 22, 209660,
                            96260, 10000, 20000, 95250, 190500, '', shape
                            )
    @obj.__send__(:write_a_xfrm, 100, 200, 10, 20, shape)

    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(expected_str)
    assert_equal(expected, result)
  end

  def expected_str
<<EOS
<a:xfrm rot="5400000" flipH="1" flipV="1">
<a:off x="100" y="200"/>
<a:ext cx="10" cy="20"/>
</a:xfrm>
EOS
  end
end
