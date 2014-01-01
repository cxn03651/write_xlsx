# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteHyperlink < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_attributes
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory('')
    result    = hyperlink.attributes(0, 0, 1)
    expected  = [ ['ref', 'A1'], ['r:id', 'rId1']]
    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_sheet2
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory('internal:Sheet2!A1', 'Sheet2!A1')
    result = hyperlink.attributes(0, 0)
    expected = [%w(ref A1), %w(location Sheet2!A1), %w(display Sheet2!A1)]
    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_quoted_sheet
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory("internal:'Data Sheet'!D5", "'Data Sheet'!D5")
    result = hyperlink.attributes(4, 0)
    expected = [%w(ref A5), ["location", "'Data Sheet'!D5"], ["display", "'Data Sheet'!D5"]]
    assert_equal(expected, result)
  end

  def test_write_hyperlink_internal_tooltip
    hyperlink = Writexlsx::Worksheet::Hyperlink.factory('internal:Sheet2!A1', 'Sheet2!A1', 'Screen Tip 1')
    result = hyperlink.attributes(17, 0)
    expected = [%w(ref A18), %w(location Sheet2!A1), ["tooltip", "Screen Tip 1"], %w(display Sheet2!A1)]
    assert_equal(expected, result)
  end
end
