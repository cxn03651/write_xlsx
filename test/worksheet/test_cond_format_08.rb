# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestCondFormat08 < Test::Unit::TestCase
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
     'yesterday', 'today', 'tomorrow',
     'last 7 days', 'last week', 'this week', 'next week',
     'last month', 'this month', 'next month'
    ].each do |criteria|
      @worksheet.conditional_formatting('A1:A4',
                                        {
                                          :type     => 'time_period',
                                          :criteria => criteria,
                                          :format   => nil
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
    <cfRule type="timePeriod" priority="1" timePeriod="yesterday">
      <formula>FLOOR(A1,1)=TODAY()-1</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="2" timePeriod="today">
      <formula>FLOOR(A1,1)=TODAY()</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="3" timePeriod="tomorrow">
      <formula>FLOOR(A1,1)=TODAY()+1</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="4" timePeriod="last7Days">
      <formula>AND(TODAY()-FLOOR(A1,1)&lt;=6,FLOOR(A1,1)&lt;=TODAY())</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="5" timePeriod="lastWeek">
      <formula>AND(TODAY()-ROUNDDOWN(A1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(A1,0)&lt;(WEEKDAY(TODAY())+7))</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="6" timePeriod="thisWeek">
      <formula>AND(TODAY()-ROUNDDOWN(A1,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(A1,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="7" timePeriod="nextWeek">
      <formula>AND(ROUNDDOWN(A1,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(A1,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="8" timePeriod="lastMonth">
      <formula>AND(MONTH(A1)=MONTH(TODAY())-1,OR(YEAR(A1)=YEAR(TODAY()),AND(MONTH(A1)=1,YEAR(A1)=YEAR(TODAY())-1)))</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="9" timePeriod="thisMonth">
      <formula>AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))</formula>
    </cfRule>
    <cfRule type="timePeriod" priority="10" timePeriod="nextMonth">
      <formula>AND(MONTH(A1)=MONTH(TODAY())+1,OR(YEAR(A1)=YEAR(TODAY()),AND(MONTH(A1)=12,YEAR(A1)=YEAR(TODAY())+1)))</formula>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
  end
end
