# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteColumn < Test::Unit::TestCase
  def test_write_auto_column
    vml = Writexlsx::Package::VML.new
    vml.__send__('write_column', 2)
    result = vml.instance_variable_get(:@writer).string
    expected = '<x:Column>2</x:Column>'
    assert_equal(expected, result)
  end
end
