# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWritePageSetUpPr < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_page_set_up_pr
    @worksheet.instance_variable_get(:@page_setup).fit_page = true
    @worksheet.__send__('write_page_set_up_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<pageSetUpPr fitToPage="1"/>'
    assert_equal(expected, result)
  end
end
