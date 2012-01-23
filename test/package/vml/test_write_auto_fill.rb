# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteAutoFill < Test::Unit::TestCase
  def test_write_auto_fill
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_auto_fill')
    result = vml.instance_variable_get(:@writer).string
    expected = '<x:AutoFill>False</x:AutoFill>'
    assert_equal(expected, result)
  end
end
