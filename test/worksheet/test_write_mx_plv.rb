# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx/workbook'
require 'write_xlsx/worksheet'
require 'stringio'

class TestWriteMxPlv < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_mx_plv
    @worksheet.__send__('write_mx_plv')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<mx:PLV Mode="1" OnePage="0" WScale="0"/>'
    assert_equal(expected, result)
  end
end
