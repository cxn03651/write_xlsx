# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteAnchor < Test::Unit::TestCase
  def test_write_anchor
    vml = Writexlsx::Package::VML.new
    vml.__send__('write_anchor', [ 2, 0, 15, 10, 4, 4, 15, 4 ])
    result = vml.instance_variable_get(:@writer).string
    expected = '<x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor>'
    assert_equal(expected, result)
  end
end
