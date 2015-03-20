# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/chart'

class TestWriteStyle < Test::Unit::TestCase
  def setup
    @chart = Writexlsx::Chart.new('Bar')
  end

  def test_write_style_1
    expected = '<c:style val="1"/>'
    @chart.set_style(1)
    @chart.__send__('write_style')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end

  def test_write_style_with_default_style_not_written
    expected = ''
    @chart.set_style(2)
    @chart.__send__('write_style')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end

  def test_write_style_with_outside_range
    expected = ''
    @chart.set_style(-1)
    @chart.__send__('write_style')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end

  def test_write_style_with_outside_range_49
    expected = ''
    @chart.set_style(49)
    @chart.__send__('write_style')
    result = @chart.instance_variable_get(:@writer).string
    assert_equal(expected, result)
  end
end
