# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWorksheetWritePageMargins < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_page_margins
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_05
    @worksheet.margins = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_LR_05
    @worksheet.margins_left_right = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_TB_05
    @worksheet.margins_top_bottom = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.7" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_left_05
    @worksheet.margin_left = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.5" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_right_05
    @worksheet.margin_right = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_top_05
    @worksheet.margin_top = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.7" top="0.5" bottom="0.75" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_bottom_05
    @worksheet.margin_bottom = 0.5
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.5" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_header_05
    @worksheet.set_header('', 0.5)
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.5" footer="0.3"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_footer_05
    @worksheet.set_footer('', 0.5)
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.5"/>'
    assert_equal(expected, result)
  end

  def test_write_page_margins_with_white_space
    @worksheet.margins = " 0.5\n"
    @worksheet.__send__('write_page_margins')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
    assert_equal(expected, result)
  end
end
