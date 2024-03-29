# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx/workbook'
require 'stringio'

class TestWorkbook03 < Minitest::Test
  def test_workbook_03
    workbook = Writexlsx::Workbook.new(StringIO.new)
    workbook.add_worksheet('Non Default Name')
    workbook.add_worksheet('Another Name')
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
    <sheet name="Non Default Name" sheetId="1" r:id="rId1"/>
    <sheet name="Another Name" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="124519" fullCalcOnLoad="1"/>
</workbook>
EOS
                                )

    assert_equal(expected, result)
  end
end
