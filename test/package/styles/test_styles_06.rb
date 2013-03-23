# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/styles'
require 'stringio'

class TestStyles06 < Test::Unit::TestCase
  def test_styles_06
    workbook = WriteXLSX.new(StringIO.new)

    format1 = workbook.add_format(
      :left         => 1,
      :right        => 1,
      :top          => 1,
      :bottom       => 1,
      :diag_border  => 1,
      :diag_type    => 3,
      :left_color   => 'red',
      :right_color  => 'red',
      :top_color    => 'red',
      :bottom_color => 'red',
      :diag_color   => 'red'
    )

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
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="2">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border diagonalUp="1" diagonalDown="1">
      <left style="thin">
        <color rgb="FFFF0000"/>
      </left>
      <right style="thin">
        <color rgb="FFFF0000"/>
      </right>
      <top style="thin">
        <color rgb="FFFF0000"/>
      </top>
      <bottom style="thin">
        <color rgb="FFFF0000"/>
      </bottom>
      <diagonal style="thin">
        <color rgb="FFFF0000"/>
      </diagonal>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
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
