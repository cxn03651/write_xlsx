# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'write_xlsx/shape'
require 'write_xlsx/drawing'
require 'stringio'

class DrawingShape04 < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_drawing_shape_04
    shape = Writexlsx::Shape.new(text: 'test', id: 1000)

    # Mock up the color palette.
    shape.palette[0] = [0x00, 0x00, 0x00, 0x00]
    shape.palette[7] = [0x00, 0x00, 0x00, 0x00]

    @obj = Writexlsx::Drawings.new
    @obj.embedded = 2
    dimensions = [
      4, 8, 209550, 95250, 12, 22, 209660, 96260, 10000, 20000
    ]
    drawing = Writexlsx::Drawing.new(
      3, dimensions, 95250, 190500, shape, 1
    )
    @obj.add_drawing_object(drawing)
    @obj.assemble_xml_file

    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(expected_str)

    assert_equal(expected, result)
  end

  def expected_str
    <<EOS
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
        <xdr:sp macro="" textlink="">
            <xdr:nvSpPr>
                <xdr:cNvPr id="1000" name="rect 1"/>
                <xdr:cNvSpPr>
                    <a:spLocks noChangeArrowheads="1"/>
                </xdr:cNvSpPr>
            </xdr:nvSpPr>
            <xdr:spPr bwMode="auto">
                <a:xfrm>
                    <a:off x="10000" y="20000"/>
                    <a:ext cx="95250" cy="190500"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
                <a:noFill/>
                <a:ln w="9525">
                    <a:solidFill>
                        <a:srgbClr val="000000"/>
                    </a:solidFill>
                    <a:miter lim="800000"/>
                    <a:headEnd/>
                    <a:tailEnd/>
                </a:ln>
            </xdr:spPr>
            <xdr:txBody>
                <a:bodyPr vertOverflow="clip" wrap="square" lIns="27432" tIns="22860" rIns="27432" bIns="22860" anchor="ctr" upright="1"/>
                <a:lstStyle/>
                <a:p>
                    <a:pPr algn="ctr" rtl="0">
                        <a:defRPr sz="1000"/>
                    </a:pPr>
                    <a:r>
                        <a:rPr lang="en-US" sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0">
                            <a:solidFill>
                                <a:srgbClr val="000000"/>
                            </a:solidFill>
                            <a:latin typeface="Calibri"/>
                            <a:cs typeface="Calibri"/>
                        </a:rPr>
                        <a:t>test</a:t>
                    </a:r>
                </a:p>
            </xdr:txBody>
        </xdr:sp>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>
EOS
  end
end
