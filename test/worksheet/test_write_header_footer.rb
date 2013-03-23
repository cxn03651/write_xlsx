# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteHeaderFooter < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
    @writer    = @worksheet.instance_variable_get(:@writer)
  end

  def test_write_odd_header
    @worksheet.set_header('Page &P of &N')
    @worksheet.
      instance_variable_get(:@page_setup).
      __send__('write_odd_header', @writer)
    result = @writer.string
    expected = '<oddHeader>Page &amp;P of &amp;N</oddHeader>'
    assert_equal(expected, result)
  end

  def test_write_odd_footer
    @worksheet.set_footer('&F')
    @worksheet.
      instance_variable_get(:@page_setup).
      __send__('write_odd_footer', @writer)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<oddFooter>&amp;F</oddFooter>'
    assert_equal(expected, result)
  end

  def test_write_haeder_footer_only_header
    @worksheet.set_header('Page &P of &N')
    @worksheet.__send__('write_header_footer')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader></headerFooter>'
    assert_equal(expected, result)
  end

  def test_write_haeder_footer_only_footer
    @worksheet.set_footer('&F')
    @worksheet.__send__('write_header_footer')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<headerFooter><oddFooter>&amp;F</oddFooter></headerFooter>'
    assert_equal(expected, result)
  end

  def test_write_haeder_footer_both_header_and_footer
    @worksheet.set_header('Page &P of &N')
    @worksheet.set_footer('&F')
    @worksheet.__send__('write_header_footer')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<headerFooter><oddHeader>Page &amp;P of &amp;N</oddHeader><oddFooter>&amp;F</oddFooter></headerFooter>'
    assert_equal(expected, result)
  end
end
