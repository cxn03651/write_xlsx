# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/core'

class TestCore02 < Test::Unit::TestCase
  def test_assemble_xml_file
    properties = {
      :title    => 'This is an example spreadsheet',
      :subject  => 'With document properties',
      :author   => 'John McNamara',
      :manager  => 'Dr. Heinz Doofenshmirtz',
      :company  => 'of Wolves',
      :category => 'Example spreadsheets',
      :keywords => 'Sample, Example, Properties',
      :comments => 'Created with Ruby and WriteXLSX',
      :status   => 'Quo',
      :created  => Time.local(2011, 4, 6, 19, 45, 15)
    }

    @obj = Writexlsx::Package::Core.new
    @obj.set_properties(properties)
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>This is an example spreadsheet</dc:title>
  <dc:subject>With document properties</dc:subject>
  <dc:creator>John McNamara</dc:creator>
  <cp:keywords>Sample, Example, Properties</cp:keywords>
  <dc:description>Created with Ruby and WriteXLSX</dc:description>
  <cp:lastModifiedBy>John McNamara</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2011-04-06T19:45:15Z</dcterms:modified>
  <cp:category>Example spreadsheets</cp:category>
  <cp:contentStatus>Quo</cp:contentStatus>
</cp:coreProperties>
EOS
    )
    assert_equal(expected, result)
  end
end
