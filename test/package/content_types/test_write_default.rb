# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/content_types'

class TestWriteDefault < Test::Unit::TestCase
  def test_assemble_xml_file
    @obj = Writexlsx::Package::ContentTypes.new(nil)
    @obj.__send__('write_default', 'xml', 'application/xml')
    result = @obj.xml_str
    expected = '<Default Extension="xml" ContentType="application/xml"/>'
    assert_equal(expected, result)
  end
end
