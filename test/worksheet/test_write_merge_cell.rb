# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteMergeCell < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_merge_cell
    @worksheet.__send__('write_merge_cell', [ 2, 1, 2, 2 ])
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<mergeCell ref="B3:C3"/>'
    assert_equal(expected, result)
  end
end
