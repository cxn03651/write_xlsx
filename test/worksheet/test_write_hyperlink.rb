# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteHyperlink < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_hyperlink_external
    @worksheet.__send__('write_hyperlink_external', 0, 0, 1)
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<hyperlink ref="A1" r:id="rId1" />'
    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_sheet2
    @worksheet.__send__('write_hyperlink_internal', 0, 0, 'Sheet2!A1', 'Sheet2!A1')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<hyperlink ref="A1" location="Sheet2!A1" display="Sheet2!A1" />'
    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_quated_sheet
    @worksheet.__send__('write_hyperlink_internal', 4, 0, "'Data Sheet'!D5", "'Data Sheet'!D5")
    result = @worksheet.instance_variable_get(:@writer).string
    expected = %q{<hyperlink ref="A5" location="'Data Sheet'!D5" display="'Data Sheet'!D5" />}
    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_tooltip
    @worksheet.__send__('write_hyperlink_internal', 17, 0, 'Sheet2!A1', 'Sheet2!A1', 'Screen Tip 1')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<hyperlink ref="A18" location="Sheet2!A1" tooltip="Screen Tip 1" display="Sheet2!A1" />'
    assert_equal(expected, result)
  end
end
