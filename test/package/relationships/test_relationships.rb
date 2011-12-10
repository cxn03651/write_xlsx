# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/relationships'

class TestRelationships < Test::Unit::TestCase
  def test_assemble_xml_file
    @obj = Writexlsx::Package::Relationships.new
    @obj.add_document_relationship('/worksheet',     'worksheets/sheet1.xml')
    @obj.add_document_relationship('/theme',         'theme/theme1.xml')
    @obj.add_document_relationship('/styles',        'styles.xml')
    @obj.add_document_relationship('/sharedStrings', 'sharedStrings.xml')
    @obj.add_document_relationship('/calcChain',     'calcChain.xml')
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>
EOS
    )
    assert_equal(expected, result)
  end
end
