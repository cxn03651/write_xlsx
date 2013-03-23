# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWorksheetWritePageSetup < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_page_setup
    assert @worksheet
  end

  def test_write_page_setup2
    @worksheet.__send__('write_page_setup')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = ''
    assert_equal(expected, result)
  end

  def test_write_page_setup_with_set_landscape
    @worksheet.set_landscape
    @worksheet.__send__('write_page_setup')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageSetup orientation="landscape"/>'
    assert_equal(expected, result)
  end

  def test_write_page_setup_with_set_portrait
    @worksheet.set_portrait
    @worksheet.__send__('write_page_setup')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageSetup orientation="portrait"/>'
    assert_equal(expected, result)
  end

  def test_write_page_setup_with_set_paper
    @worksheet.paper = 9
    @worksheet.__send__('write_page_setup')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageSetup paperSize="9" orientation="portrait"/>'
    assert_equal(expected, result)
  end

  def test_write_page_setup_with_print_across
    @worksheet.print_across
    @worksheet.__send__('write_page_setup')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageSetup pageOrder="overThenDown" orientation="portrait"/>'
    assert_equal(expected, result)
  end
end
