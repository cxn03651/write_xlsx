# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/styles'
require 'stringio'

class TestStyles07 < Test::Unit::TestCase
  def test_styles_07
    workbook = WriteXLSX.new(StringIO.new)

    format1 = workbook.add_format(:pattern => 1,  :bg_color => 'red')
    format2 = workbook.add_format(:pattern => 11, :bg_color => 'red')
    format3 = workbook.add_format(:pattern => 11, :bg_color => 'red', :fg_color => 'yellow')
    format4 = workbook.add_format(:pattern => 1,  :bg_color => 'red', :fg_color => 'red')

    workbook.__send__('set_default_xf_indices')
    workbook.__send__('prepare_format_properties')

    @style = Writexlsx::Package::Styles.new
    @style.set_style_properties(*workbook.style_properties)
    @style.assemble_xml_file
    result = got_to_array(@style.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="6">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFF0000"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="lightHorizontal">
        <bgColor rgb="FFFF0000"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="lightHorizontal">
        <fgColor rgb="FFFFFF00"/>
        <bgColor rgb="FFFF0000"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFF0000"/>
        <bgColor rgb="FFFF0000"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
EOS
      )
      assert_equal(expected, result)
    end
end
