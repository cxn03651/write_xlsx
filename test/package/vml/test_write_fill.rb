# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteFill < Test::Unit::TestCase
  def test_write_fill
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_fill')
    result = vml.instance_variable_get(:@writer).string
    expected = '<v:fill color2="#ffffe1"/>'
    assert_equal(expected, result)
  end
end
