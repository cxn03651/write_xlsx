# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteArrayFormula01 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_array_formula_does_not_change_formula
    formula = '{=TREND(C1:C3,B1:B3)}'.freeze
    assert_nothing_raised do
      @worksheet.write_array_formula('A1:A3', formula)
    end
  end

  def test_write_array_formula
    format = nil
    @worksheet.select
    @worksheet.write('B1', [ [ 500, 10 ], [ 300, 15 ] ])
    @worksheet.write('B5', [ [ 1, 2, 3 ], [ 20234, 21003, 10000 ] ])

    @worksheet.write('A1', '{=SUM(B1:C1*B2:C2)}', format, 9500)
    @worksheet.write_array_formula('A2:A2', '{=SUM(B1:C1*B2:C2)}',   format, 9500)
    @worksheet.write_array_formula('A5:A7', '{=TREND(C5:C7,B5:B7)}', format, 22196)
    @worksheet.__send__('assemble_xml_file')
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:C7"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1">
        <f t="array" ref="A1">SUM(B1:C1*B2:C2)</f>
        <v>9500</v>
      </c>
      <c r="B1">
        <v>500</v>
      </c>
      <c r="C1">
        <v>300</v>
      </c>
    </row>
    <row r="2" spans="1:3">
      <c r="A2">
        <f t="array" ref="A2">SUM(B1:C1*B2:C2)</f>
        <v>9500</v>
      </c>
      <c r="B2">
        <v>10</v>
      </c>
      <c r="C2">
        <v>15</v>
      </c>
    </row>
    <row r="5" spans="1:3">
      <c r="A5">
        <f t="array" ref="A5:A7">TREND(C5:C7,B5:B7)</f>
        <v>22196</v>
      </c>
      <c r="B5">
        <v>1</v>
      </c>
      <c r="C5">
        <v>20234</v>
      </c>
    </row>
    <row r="6" spans="1:3">
      <c r="A6">
        <v>0</v>
      </c>
      <c r="B6">
        <v>2</v>
      </c>
      <c r="C6">
        <v>21003</v>
      </c>
    </row>
    <row r="7" spans="1:3">
      <c r="A7">
        <v>0</v>
      </c>
      <c r="B7">
        <v>3</v>
      </c>
      <c r="C7">
        <v>10000</v>
      </c>
    </row>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
)
    assert_equal(expected, result)
  end
end
