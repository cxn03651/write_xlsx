# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteShapelayout < Test::Unit::TestCase
  def test_write_shapelayout
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_shapelayout', 1)
    result = vml.instance_variable_get(:@writer).string
    expected = '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
    assert_equal(expected, result)
  end
end
