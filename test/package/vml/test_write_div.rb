# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteDiv < Test::Unit::TestCase
  def test_write_div
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_div', 'left')
    result = vml.instance_variable_get(:@writer).string
    expected = '<div style="text-align:left"></div>'
    assert_equal(expected, result)
  end
end
