# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteWorksheetAttributes < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_worksheet_attributes
    result = @worksheet.__send__('write_worksheet_attributes')
    expected = [
                ['xmlns', "http://schemas.openxmlformats.org/spreadsheetml/2006/main"],
                ['xmlns:r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships"]
               ]
    assert_equal(expected, result)
  end
end
