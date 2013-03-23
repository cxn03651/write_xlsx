# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWorksheet03 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_assemble_xml_file_set_column
    format = Writexlsx::Format.new(Writexlsx::Formats.new, :xf_index => 1, :bold => 1)
    @worksheet.select
    @worksheet.set_column('B:D', 5)
    @worksheet.set_column('F:F', 8, nil, 1)
    @worksheet.set_column('H:H', nil, format)
    @worksheet.set_column('J:J', 2)
    @worksheet.set_column('L:L', nil, nil, 1)
    @worksheet.assemble_xml_file
    result = got_to_array(@worksheet.instance_variable_get(:@writer).string)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="F1:H1"/>
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="2" max="4" width="5.7109375" customWidth="1"/>
    <col min="6" max="6" width="8.7109375" hidden="1" customWidth="1"/>
    <col min="8" max="8" width="9.140625" style="1"/>
    <col min="10" max="10" width="2.7109375" customWidth="1"/>
    <col min="12" max="12" width="0" hidden="1" customWidth="1"/>
  </cols>
  <sheetData/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
EOS
    )
    assert_equal(expected, result)
  end
end
