# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestPixelsToRowCol < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def width_to_pixels(width)
    max_digit_width = 7.0
    padding         = 5

    if width < 1
      ((width * (max_digit_width + padding)) + 0.5).to_i
    else
      ((width * max_digit_width) + 0.5).to_i + padding
    end
  end

  def height_to_pixels(height)
    (4.0 * height / 3).to_i
  end

  def test_pixel_to_width
    1791.times do |pixels|
      caption  = "\tWorksheet: pixcel_to_width(#{pixels})"
      expected = pixels
      result   = width_to_pixels(@worksheet.__send__(:pixels_to_width, pixels))

      assert_equal(expected, result, caption)
    end
  end

  def test_pixel_to_height
    546.times do |pixels|
      caption  = "\tWorksheet: pixcel_to_height(#{pixels})"
      expected = pixels
      result   = height_to_pixels(@worksheet.__send__(:pixels_to_height, pixels))

      assert_equal(expected, result, caption)
    end
  end
end
