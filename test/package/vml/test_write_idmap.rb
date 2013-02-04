# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteIdmap < Test::Unit::TestCase
  def test_write_idmap
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_idmap', 1)
    result = vml.instance_variable_get(:@writer).string
    expected = '<o:idmap v:ext="edit" data="1"/>'
    assert_equal(expected, result)
  end
end
