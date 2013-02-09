# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestCondFormat20 < Test::Unit::TestCase
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

    @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'text',
                                        :criteria => 'begins with',
                                        :value    => 'b',
                                        :format   => nil
                                      }
                                      )

    @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'text',
                                        :criteria => 'begins with',
                                        :value    => 'bc',
                                        :format   => nil
                                      }
                                      )

     @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'text',
                                        :criteria => 'ends with',
                                        :value    => 'z',
                                        :format   => nil
                                      }
                                      )

     @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'text',
                                        :criteria => 'ends with',
                                        :value    => 'yz',
                                        :format   => nil
                                      }
                                      )

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
    <cfRule type="beginsWith" priority="1" operator="beginsWith" text="b">
      <formula>LEFT(A1,1)="b"</formula>
    </cfRule>
    <cfRule type="beginsWith" priority="2" operator="beginsWith" text="bc">
      <formula>LEFT(A1,2)="bc"</formula>
    </cfRule>
    <cfRule type="endsWith" priority="3" operator="endsWith" text="z">
      <formula>RIGHT(A1,1)="z"</formula>
    </cfRule>
    <cfRule type="endsWith" priority="4" operator="endsWith" text="yz">
      <formula>RIGHT(A1,2)="yz"</formula>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
  end
end
