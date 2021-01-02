# -*- coding: utf-8 -*-
require 'helper'

class TestWorksheetDataBar08 < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  ###############################################################################
  #
  # Tests for WriteXLSX::Worksheet methods.
  #
  #
  def test_data_bar_08
    @worksheet.select

    @worksheet.conditional_formatting(
      'A1',
      :type     => 'data_bar',
      :bar_only => 1
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
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
  <conditionalFormatting sqref="A1">
    <cfRule type="dataBar" priority="1">
      <dataBar showValue="0">
        <cfvo type="min" val="0"/>
        <cfvo type="max" val="0"/>
        <color rgb="FF638EC6"/>
      </dataBar>
    </cfRule>
  </conditionalFormatting>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
  end
end
