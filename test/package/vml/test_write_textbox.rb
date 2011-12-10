# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteTextbox < Test::Unit::TestCase
  def test_write_textbox
    vml = Writexlsx::Package::VML.new
    vml.__send__('write_textbox')
    result = vml.instance_variable_get(:@writer).string
    expected = '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
    assert_equal(expected, result)
  end
end
