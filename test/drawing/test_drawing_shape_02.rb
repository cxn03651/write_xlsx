# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/shape'
require 'write_xlsx/drawing'

class DrawingShape02 < Test::Unit::TestCase
  def test_drawing_shape_02
    shape = Writexlsx::Shape.new
    # Set shape properties via []= method
    shape.id = 1000
    shape.name = 'Connector 1'

    # Set bulk shape properties via set_properties method
    shape.set_properties(:type => 'straightConnector1', :connect => 1)

    @obj = Writexlsx::Drawing.new
    @obj.embedded = 1
    @obj.add_drawing_object(
                               3, 4, 8, 209550, 95250, 12, 22, 209660,
                               96260, 10000, 20000, 95250, 190500, 'Connector 1', shape
                               )
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <xdr:twoCellAnchor>
        <xdr:from>
            <xdr:col>4</xdr:col>
            <xdr:colOff>209550</xdr:colOff>
            <xdr:row>8</xdr:row>
            <xdr:rowOff>95250</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>12</xdr:col>
            <xdr:colOff>209660</xdr:colOff>
            <xdr:row>22</xdr:row>
            <xdr:rowOff>96260</xdr:rowOff>
        </xdr:to>
        <xdr:cxnSp macro="">
            <xdr:nvCxnSpPr>
                <xdr:cNvPr id="1000" name="Connector 1"/>
                <xdr:cNvCxnSpPr>
                    <a:cxnSpLocks noChangeShapeType="1"/>
                </xdr:cNvCxnSpPr>
            </xdr:nvCxnSpPr>
            <xdr:spPr bwMode="auto">
                <a:xfrm>
                    <a:off x="10000" y="20000"/>
                    <a:ext cx="95250" cy="190500"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:noFill/>
                <a:ln w="9525">
                    <a:solidFill>
                        <a:srgbClr val="000000"/>
                    </a:solidFill>
                    <a:round/>
                    <a:headEnd/>
                    <a:tailEnd/>
                </a:ln>
            </xdr:spPr>
        </xdr:cxnSp>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>
EOS
    )
    assert_equal(expected, result)
  end
end
