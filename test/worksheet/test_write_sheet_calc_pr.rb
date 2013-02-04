# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetCalcPr < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_calc_pr
    @worksheet.__send__('write_sheet_calc_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetCalcPr fullCalcOnLoad="1"/>'
    assert_equal(expected, result)
  end
end
