# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteDimension < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_dimension_with_no_dimension_set
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<dimension ref="A1"/>'
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set
    cell = 'A1'
    @worksheet.write(cell, 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_big_row
    cell = 'A1048576'
    @worksheet.write(cell, 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_big_column
    cell = 'XFD1'
    @worksheet.write(cell, 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_big_row_and_column
    cell = 'XFD1048576'
    @worksheet.write(cell, 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_narrow_range
    cell = 'A1:B2'
    @worksheet.write('A1', 'some string')
    @worksheet.write('B2', 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_narrow_range_2
    cell = 'A1:B2'
    @worksheet.write('B2', 'some string')
    @worksheet.write('A1', 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_wide_range
    cell = 'B2:H11'
    @worksheet.write('B2', 'some string')
    @worksheet.write('H11', 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end

  def test_write_dimension_with_dimension_set_very_wide_range
    cell = 'A1:XFD1048576'
    @worksheet.write('A1', 'some string')
    @worksheet.write('XFD1048576', 'some string')
    @worksheet.__send__('write_dimension')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %Q!<dimension ref="#{cell}"/>!
    assert_equal(expected, result)
  end
end
