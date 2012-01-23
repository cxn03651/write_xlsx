# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteShadow < Test::Unit::TestCase
  def test_write_shadow
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_shadow')
    result = vml.instance_variable_get(:@writer).string
    expected = '<v:shadow on="t" color="black" obscured="t" />'
    assert_equal(expected, result)
  end
end
