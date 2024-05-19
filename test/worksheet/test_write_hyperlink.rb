# -*- coding: utf-8 -*-

require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteHyperlink < Minitest::Test
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_hyperlink_internal_sheet2
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory('internal:Sheet2!A1', 'Sheet2!A1')
    result = hyperlink.attributes(0, 0)
    expected = [%w[ref A1], %w[location Sheet2!A1], %w[display Sheet2!A1]]

    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_quoted_sheet
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory("internal:'Data Sheet'!D5", "'Data Sheet'!D5")
    result = hyperlink.attributes(4, 0)
    expected = [%w[ref A5], ["location", "'Data Sheet'!D5"], ["display", "'Data Sheet'!D5"]]

    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_tooltip
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory('internal:Sheet2!A1', 'Sheet2!A1', 'Screen Tip 1')
    result = hyperlink.attributes(17, 0)
    expected = [%w[ref A18], %w[location Sheet2!A1], ["tooltip", "Screen Tip 1"], %w[display Sheet2!A1]]

    assert_equal(expected, result)
  end

  def test_raise_if_url_size_is_longer_than_default_limit
    base_url = "https://www.ruby-lang.org/"

    exceed_url = base_url + ("a" * (@workbook.max_url_length - base_url.size + 1))
    e = assert_raises RuntimeError do
      @worksheet.write_url('A1', exceed_url)
    end

    assert_match(/characters since it exceeds Excel's limit for URLS\. See LIMITATIONS section of the WriteXLSX documentation\./, e.message)
  end

  def test_raise_if_url_size_is_longer_than_specified_limit
    base_url = "https://www.ruby-lang.org/"
    max_url_length = 255

    @workbook  = WriteXLSX.new(StringIO.new, max_url_length: max_url_length)
    @worksheet = @workbook.add_worksheet

    exceed_url = base_url + ("a" * (max_url_length - base_url.size + 1))
    e = assert_raises RuntimeError do
      @worksheet.write_url('A1', exceed_url)
    end

    assert_match(/characters since it exceeds Excel's limit for URLS\. See LIMITATIONS section of the WriteXLSX documentation\./, e.message)
  end

  def assert_nothing_raised
    yield.tap { assert(true) }
  rescue StandardError => e
    raise Minitest::UnexpectedError.new(e)
  end

  def test_nothing_raise_when_nil_format_and_string_value1
    @worksheet.write_url('A1', 'http://www.ruby-lang.org', nil, 'Ruby')

    assert_nothing_raised do
      @workbook.close
    end
  end
end
