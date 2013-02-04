# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetFormatPr < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_format_pr
    @worksheet.__send__('write_sheet_format_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetFormatPr defaultRowHeight="15"/>'
    assert_equal(expected, result)
  end
end
