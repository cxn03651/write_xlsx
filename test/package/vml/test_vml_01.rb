# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteTextbox < Test::Unit::TestCase
  def test_write_textbox
    workbook  = WriteXLSX.new(StringIO.new)
    worksheet = workbook.add_worksheet
    comment = Writexlsx::Package::Comment.new(workbook, worksheet, 1, 1, 'Some text', :author => 'John', :visible => nil)
    comment.instance_variable_set(:@vertices, [2, 0, 15, 10, 4, 4, 15, 4, 143, 10, 128, 74])
    vml = Writexlsx::Package::VML.new
    vml.assemble_xml_file(1, 1024, [ comment ])
    result = got_to_array(vml.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout>
 <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path gradientshapeok="t" o:connecttype="rect"/>
 </v:shapetype>
 <v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:107.25pt;margin-top:7.5pt;width:96pt;height:55.5pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">
  <v:fill color2="#ffffe1"/>
  <v:shadow on="t" color="black" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style="mso-direction-alt:auto">
   <div style="text-align:left">
   </div>
  </v:textbox>
  <x:ClientData ObjectType="Note">
   <x:MoveWithCells/>
   <x:SizeWithCells/>
   <x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:Row>1</x:Row>
   <x:Column>1</x:Column>
  </x:ClientData>
 </v:shape>
 </xml>
EOS
    )
    assert_equal(expected, result)
  end
end
