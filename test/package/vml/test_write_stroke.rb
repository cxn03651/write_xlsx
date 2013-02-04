# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteStroke < Test::Unit::TestCase
  def test_write_stroke
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_stroke')
    result = vml.instance_variable_get(:@writer).string
    expected = '<v:stroke joinstyle="miter"/>'
    assert_equal(expected, result)
  end
end
