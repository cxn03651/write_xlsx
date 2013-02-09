# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestSparkline06 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet
  end

  def test_sparkline06
    @worksheet.instance_variable_set(:@excel_version, 2010)
    @worksheet.select

    data = [-2, 2, 3, -1, 0]

    @worksheet.write('A1', data)
    @worksheet.write('A2', data)

    # Set up sparklines

    @worksheet.add_sparkline(
                             {
                               :location => ['F1',    'F2'],
                               :range    => ['A1:E1', 'A2:E2']
                             }
                             )

    # End sparklines

    @worksheet.assemble_xml_file
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)

    expected = expected_to_array(expected_xml)
    assert_equal(expected, result)
  end

  def expected_xml
    <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
  <dimension ref="A1:E2"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
  <sheetData>
    <row r="1" spans="1:5" x14ac:dyDescent="0.25">
      <c r="A1">
        <v>-2</v>
      </c>
      <c r="B1">
        <v>2</v>
      </c>
      <c r="C1">
        <v>3</v>
      </c>
      <c r="D1">
        <v>-1</v>
      </c>
      <c r="E1">
        <v>0</v>
      </c>
    </row>
    <row r="2" spans="1:5" x14ac:dyDescent="0.25">
      <c r="A2">
        <v>-2</v>
      </c>
      <c r="B2">
        <v>2</v>
      </c>
      <c r="C2">
        <v>3</v>
      </c>
      <c r="D2">
        <v>-1</v>
      </c>
      <c r="E2">
        <v>0</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <extLst>
    <ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup displayEmptyCellsAs="gap">
          <x14:colorSeries theme="4" tint="-0.499984740745262"/>
          <x14:colorNegative theme="5"/>
          <x14:colorAxis rgb="FF000000"/>
          <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
          <x14:colorFirst theme="4" tint="0.39997558519241921"/>
          <x14:colorLast theme="4" tint="0.39997558519241921"/>
          <x14:colorHigh theme="4"/>
          <x14:colorLow theme="4"/>
          <x14:sparklines>
            <x14:sparkline>
              <xm:f>Sheet1!A1:E1</xm:f>
              <xm:sqref>F1</xm:sqref>
            </x14:sparkline>
            <x14:sparkline>
              <xm:f>Sheet1!A2:E2</xm:f>
              <xm:sqref>F2</xm:sqref>
            </x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>
EOS
  end
end
