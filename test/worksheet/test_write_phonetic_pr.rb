# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWritePhoneticPr < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_phonetic_pr
    @worksheet.__send__('write_phonetic_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<phoneticPr fontId="0" type="noConversion"/>'
    assert_equal(expected, result)
  end
end
