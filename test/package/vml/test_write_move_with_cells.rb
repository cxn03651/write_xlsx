# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/package/vml'

class TestWriteMoveWithCells < Test::Unit::TestCase
  def test_write_move_with_cells
    vml = Writexlsx::Package::VML.new
    vml.__send__('write_move_with_cells')
    result = vml.instance_variable_get(:@writer).string
    expected = '<x:MoveWithCells />'
    assert_equal(expected, result)
  end
end
