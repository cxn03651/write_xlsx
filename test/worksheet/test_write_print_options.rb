# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWritePrintOptions < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_print_options
    @worksheet.__send__('write_print_options')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = ''
    assert_equal(expected, result)
  end

  def test_write_print_options_center_horizontally
    @worksheet.center_horizontally
    @worksheet.__send__('write_print_options')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<printOptions horizontalCentered="1"/>'
    assert_equal(expected, result)
  end

  def test_write_print_options_center_vertically
    @worksheet.center_vertically
    @worksheet.__send__('write_print_options')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<printOptions verticalCentered="1"/>'
    assert_equal(expected, result)
  end

  def test_write_print_options_center_horizontally_and_vertically
    @worksheet.center_horizontally
    @worksheet.center_vertically
    @worksheet.__send__('write_print_options')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<printOptions horizontalCentered="1" verticalCentered="1"/>'
    assert_equal(expected, result)
  end

  def test_write_print_options_hide_gridlines
    @worksheet.hide_gridlines
    @worksheet.__send__('write_print_options')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = ''
    assert_equal(expected, result)
  end

  def test_write_print_options_hide_gridlines_false
    @worksheet.hide_gridlines(false)
    @worksheet.__send__('write_print_options')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<printOptions gridLines="1"/>'
    assert_equal(expected, result)
  end
=begin
  def test_write_print_options_1_hide_gridlines
    @worksheet.hide_gridlines
    @worksheet.__send__('write_print_options', 1)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = ''
    assert_equal(expected, result)
  end

  def test_write_print_options_2_hide_gridlines_false
    @worksheet.hide_gridlines(false)
    @worksheet.__send__('write_print_options', 2)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<printOptions gridLines="1"/>'
    assert_equal(expected, result)
  end
=end
end
