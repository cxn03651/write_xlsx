# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetData < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_data
    @worksheet.__send__('write_sheet_data')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetData/>'
    assert_equal(expected, result)
  end
end
