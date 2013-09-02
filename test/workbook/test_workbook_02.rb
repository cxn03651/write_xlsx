# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'stringio'

class TestWorkbook02 < Test::Unit::TestCase
  def test_workbook_01
    workbook = Writexlsx::Workbook.new(StringIO.new)
    workbook.add_worksheet
    workbook.add_worksheet
    workbook.__send__('assemble_xml_file')
    result = got_to_array(workbook.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>
EOS
    )
    assert_equal(expected, result)
  end
end
