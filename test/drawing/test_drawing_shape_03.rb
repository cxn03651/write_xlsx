# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/shape'
require 'write_xlsx/drawing'

class DrawingShape03 < Test::Unit::TestCase
  def test_drawing_shape_03
    shape = Writexlsx::Shape.new
    # Set shape properties via []= method
    shape.id          = 1000
    shape.start       = 1001
    shape.start_index = 1
    shape.end         = 1002
    shape.end_index   = 4

    @obj = Writexlsx::Drawing.new
    @obj.embedded = 1
    @obj.add_drawing_object(
                               3, 4, 8, 209550, 95250, 12, 22, 209660,
                               96260, 10000, 20000, 95250, 190500, '', shape
                               )
    @obj.__send__('write_nv_cxn_sp_pr', 1, shape)

    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<xdr:nvCxnSpPr>
<xdr:cNvPr id="1000" name="rect 1"/>
<xdr:cNvCxnSpPr>
<a:cxnSpLocks noChangeShapeType="1"/>
<a:stCxn id="1001" idx="1"/>
<a:endCxn id="1002" idx="4"/>
</xdr:cNvCxnSpPr>
</xdr:nvCxnSpPr>
EOS
    )
    assert_equal(expected, result)
  end
end
