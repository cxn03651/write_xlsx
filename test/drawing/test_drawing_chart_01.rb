# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class DrawingChart011 < Test::Unit::TestCase
  def test_drawing_chart_01
    @obj = Writexlsx::Drawing.new
    @obj.add_drawing_object(1, 4, 8, 457200, 104775, 12, 22, 152400, 180975)
    @obj.embedded = true
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>4</xdr:col>
      <xdr:colOff>457200</xdr:colOff>
      <xdr:row>8</xdr:row>
      <xdr:rowOff>104775</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>12</xdr:col>
      <xdr:colOff>152400</xdr:colOff>
      <xdr:row>22</xdr:row>
      <xdr:rowOff>180975</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
EOS
    )
    assert_equal(expected, result)
  end
end
