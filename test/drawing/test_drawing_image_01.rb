# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/drawing'

class DrawingImage011 < Test::Unit::TestCase
  def test_drawing_image_01
    @obj = Writexlsx::Drawing.new
    @obj.add_drawing_object(2, 2, 1, 0, 0, 3, 6, 533257, 190357,
                           1219200, 190500, 1142857, 1142857, 'republic.png')
    @obj.embedded = true
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>3</xdr:col>
      <xdr:colOff>533257</xdr:colOff>
      <xdr:row>6</xdr:row>
      <xdr:rowOff>190357</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1" descr="republic.png"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="1219200" y="190500"/>
          <a:ext cx="1142857" cy="1142857"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
EOS
    )
    assert_equal(expected, result)
  end
end
