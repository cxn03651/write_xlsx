# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/core'

class TestCore01 < Test::Unit::TestCase
  def test_assemble_xml_file
    properties = {
      :author   => 'A User',
      :created  => Time.local(2010, 1, 1, 0, 0, 0)
    }

    @obj = Writexlsx::Package::Core.new
    @obj.set_properties(properties)
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>A User</dc:creator>
  <cp:lastModifiedBy>A User</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2010-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2010-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>
EOS
    )
    assert_equal(expected, result)
  end
end
