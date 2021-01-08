# -*- coding: utf-8 -*-
require 'helper'

class TestComments01 < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  ###############################################################################
  #
  # Test the _assemble_xml_file() method.
  #
  def test_assemble_xml_file
    @worksheet.write_comment(
      1, 1, 'Some text',
      :author => 'John', :visible => nil, :color => 81,
      :font => 'Tahoma', :font_size => 8, :font_family => 2
    )

    comments = @worksheet.comments
    comments.assemble_xml_file
    result = got_to_array(comments.instance_variable_get(:@writer).string)

    expected = expected_to_array(expected_xml)
    assert_equal(expected, result)
  end

  def expected_xml
    <<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors>
    <author>John</author>
  </authors>
  <commentList>
    <comment ref="B2" authorId="0">
      <text>
        <r>
          <rPr>
            <sz val="8"/>
            <color indexed="81"/>
            <rFont val="Tahoma"/>
            <family val="2"/>
          </rPr>
          <t>Some text</t>
        </r>
      </text>
    </comment>
  </commentList>
</comments>
EOS
  end
end
