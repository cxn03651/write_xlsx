# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestChartsheet < Test::Unit::TestCase
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    @chartsheet = Writexlsx::Chartsheet.new(workbook, 1, '')
  end

  def test_chartsheet01
    @chartsheet.__send__(:assemble_xml_file)
    result = @chartsheet.instance_variable_get(:@writer).string
    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def expected
    <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <drawing r:id="rId1"/>
</chartsheet>
EOS
  end
end
