# -*- coding: utf-8 -*-
# frozen_string_literal: true

require 'helper'
require 'write_xlsx/package/xml_writer_simple'

class TestXMLWriterSimple < Minitest::Test
  def setup
    @obj = Writexlsx::Package::XMLWriterSimple.new
  end

  def test_xml_decl
    assert_equal(
      %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n),
      @obj.xml_decl
    )
  end

  def test_empty_tag
    assert_equal('<foo/>', @obj.empty_tag('foo'))
  end

  def test_empty_tag_with_xml_decl
    expected = <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<foo/>
EOS
    assert_equal(expected, @obj.xml_decl + @obj.empty_tag('foo') + "\n")
  end

  def test_start_end_tag
    assert_equal("<foo></foo>", @obj.start_tag('foo') + @obj.end_tag('foo'))
  end

  def test_attribute
    assert_equal(
      "<foo x=\"1&gt;2\"/>", @obj.empty_tag("foo", [['x', '1>2']])
    )
  end

  def test_character_data
    assert_equal(
      "<foo>&lt;tag&gt;&amp;amp;&lt;/tag&gt;</foo>",
      @obj.start_tag('foo') + @obj.characters("<tag>&amp;</tag>") + @obj.end_tag('foo')
    )
  end

  def test_data_element_with_empty_attr
    expected = "<foo>data</foo>"
    @obj.data_element('foo', 'data')
    result = @obj.string

    assert_equal(expected, result)
  end

  def test_data_element
    attributes = [
      ['name', '_xlnm.Print_Titles'],
      ['localSheetId', 0]
    ]
    expected =
      "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">Sheet1!$1:$1</definedName>"
    @obj.data_element('definedName', 'Sheet1!$1:$1', attributes)
    result = @obj.string

    assert_equal(expected, result)
  end
end
