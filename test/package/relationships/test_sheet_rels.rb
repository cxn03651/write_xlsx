# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/relationships'

class TestSheetRels < Test::Unit::TestCase
  def test_sheet_rels
    @obj = Writexlsx::Package::Relationships.new
    @obj.add_worksheet_relationship('/hyperlink', 'www.foo.com', 'External')
    @obj.add_worksheet_relationship('/hyperlink', 'link00.xlsx', 'External')
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="www.foo.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="link00.xlsx" TargetMode="External"/>
</Relationships>
EOS
    )
    assert_equal(expected, result)
  end
end
