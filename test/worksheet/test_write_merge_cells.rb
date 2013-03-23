# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteMergeCells < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_merge_cells_B3_C3_Foo_format
    format = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1)
    @worksheet.select
    @worksheet.merge_range('B3:C3', 'Foo', format)
    @worksheet.__send__('assemble_xml_file')
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="B3:C3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="3" spans="2:3">
      <c r="B3" s="1" t="s">
        <v>0</v>
      </c>
      <c r="C3" s="1"/>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="B3:C3"/>
  </mergeCells>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
    )
    assert_equal(expected, result)
  end

  def test_write_merge_cells_two_range
    format1 = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1)
    format2 = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 2)
    @worksheet.select
    @worksheet.merge_range('B3:C3', 'Foo', format1)
    @worksheet.merge_range('A2:D2', nil,   format2)
    @worksheet.__send__('assemble_xml_file')
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A2:D3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="2" spans="1:4">
      <c r="A2" s="2"/>
      <c r="B2" s="2"/>
      <c r="C2" s="2"/>
      <c r="D2" s="2"/>
    </row>
    <row r="3" spans="1:4">
      <c r="B3" s="1" t="s">
        <v>0</v>
      </c>
      <c r="C3" s="1"/>
    </row>
  </sheetData>
  <mergeCells count="2">
    <mergeCell ref="B3:C3"/>
    <mergeCell ref="A2:D2"/>
  </mergeCells>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
    )
    assert_equal(expected, result)
  end

  def test_write_merge_range_type
    format1 = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1)
    format2 = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 2)
    format3 = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 3)

    @worksheet.set_column('B:C', 12)
    @worksheet.instance_variable_set(:@date_1904, 0)

    @worksheet.select
    @worksheet.merge_range_type('formula',     'B14:C14', '=1+2',                 format1, 3)
    @worksheet.merge_range_type('number',      'B2:C2',   123,                    format1)
    @worksheet.merge_range_type('string',      'B4:C4',   'foo',                  format1)
    @worksheet.merge_range_type('blank',       'B6:C6',                           format1)
#    @worksheet.merge_range_type('rich_string', 'B8:C8',   'This is ', format2, 'bold', format1)
    @worksheet.merge_range_type('date_time',   'B10:C10', '2011-01-01T',          format2)
#    @worksheet.merge_range_type('url',         'B12:C12', 'http://www.perl.com/', format3)

    @worksheet.__send__('assemble_xml_file')
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="B2:C14"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="2" max="3" width="12.7109375" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="2" spans="2:3">
      <c r="B2" s="1">
        <v>123</v>
      </c>
      <c r="C2" s="1"/>
    </row>
    <row r="4" spans="2:3">
      <c r="B4" s="1" t="s">
        <v>0</v>
      </c>
      <c r="C4" s="1"/>
    </row>
    <row r="6" spans="2:3">
      <c r="B6" s="1"/>
      <c r="C6" s="1"/>
    </row>
    <row r="10" spans="2:3">
      <c r="B10" s="2">
        <v>40544</v>
      </c>
      <c r="C10" s="2"/>
    </row>
    <row r="14" spans="2:3">
      <c r="B14" s="1">
        <f>1+2</f>
        <v>3</v>
      </c>
      <c r="C14" s="1"/>
    </row>
  </sheetData>
  <mergeCells count="5">
    <mergeCell ref="B14:C14"/>
    <mergeCell ref="B2:C2"/>
    <mergeCell ref="B4:C4"/>
    <mergeCell ref="B6:C6"/>
    <mergeCell ref="B10:C10"/>
  </mergeCells>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
    )
    assert_equal(expected, result)
  end
=begin
  def test_write_merge_cells_2_1_2_2_Foo_format
    format = Writexlsx::Format.new({}, {}, :xf_index => 1)
    @worksheet.merge_range(2, 1, 2, 2, 'Foo', format)
    @worksheet.__send__('write_merge_cells')
    @worksheet.__send__('assemble_xml_file')
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="B3:C3"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>
    <row r="3" spans="2:3">
      <c r="B3" s="1" t="s">
        <v>0</v>
      </c>
      <c r="C3" s="1"/>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="B3:C3"/>
  </mergeCells>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
    )
    assert_equal(expected, result)
  end
=end
end
