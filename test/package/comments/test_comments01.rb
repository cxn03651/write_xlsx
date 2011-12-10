# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/comments'

class TestComments01 < Test::Unit::TestCase
  def test_assemble_xml_file
    @obj = Writexlsx::Package::Comments.new
    @obj.assemble_xml_file([[1, 1, 'Some text', 'John', nil, 81, [2, 0, 4, 4, 143, 10, 128, 74]]])
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
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
    )
    assert_equal(expected, result)
  end
end
