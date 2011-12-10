# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'

class TestWriteCellValue < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_cell_value
    @worksheet.__send__('write_cell_value', 1)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<v>1</v>'
    assert_equal(expected, result)
  end
end
