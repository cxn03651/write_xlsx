# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/shared_strings'

class TestSharedStrings02 < Test::Unit::TestCase
  def test_shared_strings02
    @obj = Writexlsx::Package::SharedStrings.new
    @obj.index('abcdefg')
    @obj.index('   abcdefg')
    @obj.index('abcdefg   ')
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si>
    <t>abcdefg</t>
  </si>
  <si>
    <t xml:space="preserve">   abcdefg</t>
  </si>
  <si>
    <t xml:space="preserve">abcdefg   </t>
  </si>
</sst>
EOS
    )
    assert_equal(expected, result)
  end
end
