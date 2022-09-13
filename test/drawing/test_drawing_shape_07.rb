# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'write_xlsx/shape'
require 'write_xlsx/drawing'
require 'stringio'

class DrawingShape07 < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_drawing_shape_07
    shape = Writexlsx::Shape.new
    shape.line_weight = 5
    shape.line_type   = 'lgDashDot'

    @obj = Writexlsx::Drawings.new
    @obj.embedded = 1
    dimensions = [
      4, 8, 209550, 95250, 12, 22, 209660, 96260, 10000, 20000
    ]
    drawing = Writexlsx::Drawing.new(
      3, dimensions, 95250, 190500, '', shape, 1
    )
    @obj.add_drawing_object(drawing)
    # 3, 4, 8, 209550, 95250, 12, 22, 209660,
    # 96260, 10000, 20000, 95250, 190500, '', shape
    # )
    @obj.__send__(:write_a_ln, shape)

    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(expected_str)
    assert_equal(expected, result)
  end

  def expected_str
    <<EOS
<a:ln w="47625">
<a:solidFill>
<a:srgbClr val="000000"/>
</a:solidFill>
<a:prstDash val="lgDashDot"/>
<a:miter lim="800000"/>
<a:headEnd/>
<a:tailEnd/>
</a:ln>
EOS
  end
end
