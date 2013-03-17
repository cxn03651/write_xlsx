# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/styles'
require 'stringio'

class TestStyles04 < Test::Unit::TestCase
  def test_styles_04
    workbook = WriteXLSX.new(StringIO.new)

    format1  = workbook.add_format(:top => 7)
    format2  = workbook.add_format(:top => 4)
    format3  = workbook.add_format(:top => 11)
    format4  = workbook.add_format(:top => 9)
    format5  = workbook.add_format(:top => 3)
    format6  = workbook.add_format(:top => 1)
    format7  = workbook.add_format(:top => 12)
    format8  = workbook.add_format(:top => 13)
    format9  = workbook.add_format(:top => 10)
    format10 = workbook.add_format(:top => 8)
    format11 = workbook.add_format(:top => 2)
    format12 = workbook.add_format(:top => 5)
    format13 = workbook.add_format(:top => 6)

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
  <borders count="14">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="hair">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dotted">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashDotDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="dashed">
        <color auto="1"/>
      </top>
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
      <top style="mediumDashDotDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="slantDashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashDot">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="mediumDashed">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="medium">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="thick">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
    <border>
      <left/>
      <right/>
      <top style="double">
        <color auto="1"/>
      </top>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="14">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="10" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="11" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="12" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="13" xfId="0" applyBorder="1"/>
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
