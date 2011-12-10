# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteRow < Test::Unit::TestCase
  def test_write_row
    vml = Writexlsx::Package::VML.new
    vml.__send__('write_row', 2)
    result = vml.instance_variable_get(:@writer).string
    expected = '<x:Row>2</x:Row>'
    assert_equal(expected, result)
  end
end
