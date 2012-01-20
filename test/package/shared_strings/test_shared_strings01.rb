# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/shared_strings'

class TestSharedStrings01 < Test::Unit::TestCase
  def test_shared_strings01
    @obj = Writexlsx::Package::SharedStrings.new
    @obj.index('neptune')
    @obj.index('mars')
    5.times { @obj.index('venus') }
    @obj.assemble_xml_file
    result = got_to_array(@obj.xml_str)
    expected = expected_to_array(<<EOS
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="7" uniqueCount="3">
  <si>
    <t>neptune</t>
  </si>
  <si>
    <t>mars</t>
  </si>
  <si>
    <t>venus</t>
  </si>
</sst>
EOS
    )
    assert_equal(expected, result)
  end
end
