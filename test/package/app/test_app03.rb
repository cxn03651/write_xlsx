# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/app'

class TestApp03 < Test::Unit::TestCase
  def test_assemble_xml_file
    @obj = Writexlsx::Package::App.new(nil)
    @obj.add_part_name('Sheet1')
    @obj.add_part_name('Sheet1!Print_Titles')
    @obj.add_heading_pair(['Worksheets', 1])
    @obj.add_heading_pair(['Named Ranges', 1])
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="4" baseType="variant">
      <vt:variant>
        <vt:lpstr>Worksheets</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>
      <vt:variant>
        <vt:lpstr>Named Ranges</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="2" baseType="lpstr">
      <vt:lpstr>Sheet1</vt:lpstr>
      <vt:lpstr>Sheet1!Print_Titles</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company>
  </Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>12.0000</AppVersion>
</Properties>
EOS
    )
    assert_equal(expected, result)
  end
end
