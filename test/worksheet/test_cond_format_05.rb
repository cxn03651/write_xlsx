# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestCondFormat05 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  ###############################################################################
  #
  # Test the _assemble_xml_file() method.
  #
  # Test conditional formats.
  #
  def test_conditional_formats
    @worksheet.select

    # Start test code.
    @worksheet.write('A1', 10)
    @worksheet.write('A2', 20)
    @worksheet.write('A3', 30)
    @worksheet.write('A4', 40)

    [
     'above', 'below', 'equal or above', 'equal or below',
     '1 std dev above', '1 std dev below',
     '2 std dev above', '2 std dev below',
     '3 std dev above', '3 std dev below'
    ].each do |criteria|
      @worksheet.conditional_formatting('A1:A4',
                                        {
                                          :type     => 'average',
                                          :format   => nil,
                                          :criteria => criteria
                                        }
                                        )
    end

    @worksheet.assemble_xml_file
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)

    expected = expected_to_array(expected_xml)
    assert_equal(expected, result)
  end

  def expected_xml
    <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:A4"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:1">
      <c r="A1">
        <v>10</v>
      </c>
    </row>
    <row r="2" spans="1:1">
      <c r="A2">
        <v>20</v>
      </c>
    </row>
    <row r="3" spans="1:1">
      <c r="A3">
        <v>30</v>
      </c>
    </row>
    <row r="4" spans="1:1">
      <c r="A4">
        <v>40</v>
      </c>
    </row>
  </sheetData>
  <conditionalFormatting sqref="A1:A4">
    <cfRule type="aboveAverage" priority="1"/>
    <cfRule type="aboveAverage" priority="2" aboveAverage="0"/>
    <cfRule type="aboveAverage" priority="3" equalAverage="1"/>
    <cfRule type="aboveAverage" priority="4" aboveAverage="0" equalAverage="1"/>
    <cfRule type="aboveAverage" priority="5" stdDev="1"/>
    <cfRule type="aboveAverage" priority="6" aboveAverage="0" stdDev="1"/>
    <cfRule type="aboveAverage" priority="7" stdDev="2"/>
    <cfRule type="aboveAverage" priority="8" aboveAverage="0" stdDev="2"/>
    <cfRule type="aboveAverage" priority="9" stdDev="3"/>
    <cfRule type="aboveAverage" priority="10" aboveAverage="0" stdDev="3"/>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
  end
end
