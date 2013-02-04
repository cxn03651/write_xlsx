# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteTabColor < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_tab_color
    # Mock up the color palette.
    @worksheet.instance_variable_set(:@tab_color, 0x0A)
    palette = [nil, nil, [ 0xff, 0x00, 0x00, 0x00 ]]
    @worksheet.instance_variable_set(:@palette, palette)

    @worksheet.__send__('write_tab_color')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<tabColor rgb="FFFF0000"/>'
    assert_equal(expected, result)
  end
end
