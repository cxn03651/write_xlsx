# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWorksheet04 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_assemble_xml_file_set_row
    format = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1, :bold => 1)
    @worksheet.select
    @worksheet.set_row(1, 30)
    @worksheet.set_row(3, nil, nil, 1)
    @worksheet.set_row(6, nil, format)
    @worksheet.set_row(9, 3)
    @worksheet.set_row(12, 24, nil, 1)
    @worksheet.set_row(14, 0)
    @worksheet.assemble_xml_file
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A2:A15"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="2" ht="30" customHeight="1"/>
    <row r="4" hidden="1"/>
    <row r="7" s="1" customFormat="1"/>
    <row r="10" ht="3" customHeight="1"/>
    <row r="13" ht="24" hidden="1" customHeight="1"/>
    <row r="15" hidden="1"/>
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
    )
    assert_equal(expected, result)
  end
end
