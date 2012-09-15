# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'write_xlsx/shape'
require 'write_xlsx/drawing'
require 'stringio'

class DrawingShape06 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_drawing_shape_06
    shape = Writexlsx::Shape.new
    shape.adjustments = -10, 100, 20

    @obj = Writexlsx::Drawing.new
    @obj.embedded = 1

    @obj.add_drawing_object(
                            3, 4, 8, 209550, 95250, 12, 22, 209660,
                            96260, 10000, 20000, 95250, 190500, '', shape
                            )
    @obj.__send__(:write_a_av_lst, shape)

    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(expected_str)
    assert_equal(expected, result)
  end

  def expected_str
<<EOS
<a:avLst>
<a:gd name="adj" fmla="val -10000"/>
<a:gd name="adj" fmla="val 100000"/>
<a:gd name="adj" fmla="val 20000"/>
</a:avLst>
EOS
  end
end
