# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/styles'
require 'stringio'

class TestStyles05 < Test::Unit::TestCase
  def test_styles_05
    workbook = WriteXLSX.new(StringIO.new)

    format1  = workbook.add_format(:left      => 1)
    format2  = workbook.add_format(:right     => 1)
    format3  = workbook.add_format(:top       => 1)
    format4  = workbook.add_format(:bottom    => 1)
    format5  = workbook.add_format(:diag_type => 1, :diag_border => 1)
    format6  = workbook.add_format(:diag_type => 2, :diag_border => 1)
    format7  = workbook.add_format(:diag_type => 3)

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
  <borders count="8">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left style="thin">
        <color auto="1"/>
      </left>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right style="thin">
        <color auto="1"/>
      </right>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="thin">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style="thin">
        <color auto="1"/>
      </bottom>
      <diagonal/>
    </border>
    <border diagonalUp="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color auto="1"/>
      </diagonal>
    </border>
    <border diagonalDown="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color auto="1"/>
      </diagonal>
    </border>
    <border diagonalUp="1" diagonalDown="1">
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal style="thin">
        <color auto="1"/>
      </diagonal>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
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
