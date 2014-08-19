# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteMultilineStringWithUrl < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_with_url_after_cr_string
    url_strings = %w(http:// https:// ftp:// ftps:// mailto: internal: external:)
    col = 0
    row = 0

    url_strings.each do |url_string|
      assert_nothing_raised do
        @worksheet.write(row, col, long_string(url_string))
        row += 1
      end
    end
  end

  def long_string(url_string)
    "*" * Writexlsx::Worksheet::Hyperlink::MAXIMUM_URLS_SIZE <<
      "\n" <<
      url_string
  end
end
