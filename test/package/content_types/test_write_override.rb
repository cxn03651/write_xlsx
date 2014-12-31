# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/package/content_types'

class TestWriteOverride < Test::Unit::TestCase
  def test_assemble_xml_file
    @obj = Writexlsx::Package::ContentTypes.new(nil)
    @obj.__send__('write_override', '/docProps/core.xml', 'app...')
    result = @obj.xml_str
    expected = '<Override PartName="/docProps/core.xml" ContentType="app..."/>'
    assert_equal(expected, result)
  end
end
