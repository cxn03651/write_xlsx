# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteFilters < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_filters_East
    @worksheet.__send__('write_filters', 'East')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<filters><filter val="East"/></filters>'
    assert_equal(expected, result)
  end

  def test_write_filters_East_South
    @worksheet.__send__('write_filters', 'East', 'South')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<filters><filter val="East"/><filter val="South"/></filters>'
    assert_equal(expected, result)
  end

  def test_write_filters_blanks
    @worksheet.__send__('write_filters', 'blanks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<filters blank="1"/>'
    assert_equal(expected, result)
  end
end
