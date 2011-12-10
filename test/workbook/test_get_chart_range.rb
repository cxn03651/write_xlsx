# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestGetChartRange < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
  end

  def test_get_chart_range_simple_formula
    formula = 'Sheet1!$B$1:$B$5'
    result = @workbook.__send__('get_chart_range', formula)
    expected = ['Sheet1', 0, 1, 4, 1]
    assert_equal(expected, result)
  end

  def test_get_chart_range_sheetname_with_space
    formula  = "'Sheet 1'!$B$1:$B$5"
    result = @workbook.__send__('get_chart_range', formula)
    expected = ['Sheet 1', 0, 1, 4, 1]
    assert_equal(expected, result)
  end

  def test_get_chart_range_single_cell_range
    formula  = 'Sheet1!$B$1'
    result = @workbook.__send__('get_chart_range', formula)
    expected = ['Sheet1', 0, 1, 0, 1]
    assert_equal(expected, result)
  end

  def test_get_chart_range_sheet_name_with_an_apostrophe
    formula  = "'Don''t'!$B$1:$B$5"
    result = @workbook.__send__('get_chart_range', formula)
    expected = ["Don't", 0, 1, 4, 1]
    assert_equal(expected, result)
  end

  def test_get_chart_range_sheet_name_with_exclamation_mark
    formula  = "'aa!bb'!$B$1:$B$5"
    result = @workbook.__send__('get_chart_range', formula)
    expected = ['aa!bb', 0, 1, 4, 1]
    assert_equal(expected, result)
  end

  def test_get_chart_range_sheet_name_with_invalid_range
    formula  = ''
    result = @workbook.__send__('get_chart_range', formula)
    expected = nil
    assert_equal(expected, result)
  end

  def test_get_chart_range_sheet_name_with_invalid_2d_range
    formula  = 'Sheet1!$B$1:$F$5'
    result = @workbook.__send__('get_chart_range', formula)
    expected = nil
    assert_equal(expected, result)
  end
end
