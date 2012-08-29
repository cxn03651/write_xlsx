# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestCondFormat06 < Test::Unit::TestCase
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
                                        :type     => 'top',
                                        :value    => 10,
                                        :format   => nil
                                      }
                                      )
    @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'bottom',
                                        :value    => 10,
                                        :format   => nil
                                      }
                                      )
    @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'top',
                                        :criteria => '%',
                                        :value    => 10,
                                        :format   => nil
                                      }
                                      )
    @worksheet.conditional_formatting('A1:A4',
                                      {
                                        :type     => 'bottom',
                                        :criteria => '%',
                                        :value    => 10,
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
    <cfRule type="top10" priority="1" rank="10"/>
    <cfRule type="top10" priority="2" bottom="1" rank="10"/>
    <cfRule type="top10" priority="3" percent="1" rank="10"/>
    <cfRule type="top10" priority="4" percent="1" bottom="1" rank="10"/>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
  end
end
