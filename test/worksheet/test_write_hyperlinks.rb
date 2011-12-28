# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteHyperlinks < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_hyperlinks
    @worksheet.instance_variable_set(:@hlink_refs, [[ 1, 0, 0, 1 ]])
    @worksheet.__send__('write_hyperlinks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<hyperlinks><hyperlink ref="A1" r:id="rId1" /></hyperlinks>'
    assert_equal(expected, result)
  end

  def test_write_hyperlinks_internal
    @worksheet.instance_variable_set(:@hlink_refs, [[ 2, 0, 0, 'Sheet2!A1', 'Sheet2!A1' ]])
    @worksheet.__send__('write_hyperlinks')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<hyperlinks><hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1" /></hyperlinks>'
    assert_equal(expected, result)
  end
end
