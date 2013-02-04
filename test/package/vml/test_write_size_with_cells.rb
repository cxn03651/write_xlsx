# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteSizeWithCells < Test::Unit::TestCase
  def test_write_size_with_cells
    vml = Writexlsx::Package::Vml.new
    vml.__send__('write_size_with_cells')
    result = vml.instance_variable_get(:@writer).string
    expected = '<x:SizeWithCells/>'
    assert_equal(expected, result)
  end
end
