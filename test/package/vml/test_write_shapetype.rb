# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteShapetype < Test::Unit::TestCase
  def test_write_shapetype
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_comment_shapetype')
    result = vml.instance_variable_get(:@writer).string
    expected = '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
    assert_equal(expected, result)
  end

  def test_write_shapetype2
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_button_shapetype')
    result = vml.instance_variable_get(:@writer).string
    expected = '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype>'
    assert_equal(expected, result)
  end
end
