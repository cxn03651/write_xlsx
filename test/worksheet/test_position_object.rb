# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class MockObject
  attr_reader :col, :row, :x_offset, :y_offset, :scaled_width, :scaled_height, :anchor

  def initialize(col, row, x_offset, y_offset, scaled_width, scaled_height, anchor = nil)
    @col           = col
    @row           = row
    @x_offset      = x_offset
    @y_offset      = y_offset
    @scaled_width  = scaled_width
    @scaled_height = scaled_height
    @anchor        = anchor
  end
end

class TestPositionObject < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_position_object_emus_01
    params = MockObject.new(4, 8, 0, 0, 480, 288)
    result = @worksheet.__send__("position_object_emus", params)
    expected = [4, 8, 0, 0, 11, 22, 304800, 76200, 2438400, 1524000]

    assert_equal(expected, result)
  end

  def test_position_object_emus_02
    @worksheet.set_column('L:L', 3.86)
    params = MockObject.new(4, 8, 0, 0, 480, 288)
    result = @worksheet.__send__("position_object_emus", params)
    expected = [4, 8, 0, 0, 12, 22, 0, 76200, 2438400, 1524000]

    assert_equal(expected, result)
  end

  def test_position_object_emus_03
    @worksheet.set_column('L:L', 3.86)
    @worksheet.set_row(22, 6)
    params = MockObject.new(4, 8, 0, 0, 480, 288)
    result = @worksheet.__send__("position_object_emus", params)
    expected = [4, 8, 0, 0, 12, 23, 0, 0, 2438400, 1524000]

    assert_equal(expected, result)
  end

  def test_position_object_emus_04
    params = MockObject.new(4, 8, 0, 0, 32, 32)
    result = @worksheet.__send__("position_object_emus", params)
    expected = [4, 8, 0, 0, 4, 9, 304800, 114300, 2438400, 1524000]

    assert_equal(expected, result)
  end

  def test_position_object_emus_05
    params = MockObject.new(4, 8, 2, 3, 72, 72)
    result = @worksheet.__send__("position_object_emus", params)
    expected = [4, 8, 19050, 28575, 5, 11, 95250, 142875, 2457450, 1552575]

    assert_equal(expected, result)
  end

  def test_position_object_emus_06
    params = MockObject.new(5, 1, 2, 3, 99, 69)
    result = @worksheet.__send__("position_object_emus", params)
    expected = [5, 1, 19050, 28575, 6, 4, 352425, 114300, 3067050, 219075]

    assert_equal(expected, result)
  end
end
