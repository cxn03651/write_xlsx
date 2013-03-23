# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetPr < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_pr_fit_page
    @worksheet.instance_variable_get(:@page_setup).fit_page = true
    @worksheet.__send__('write_sheet_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
    assert_equal(expected, result)
  end

  def test_write_sheet_pr_tab_color
    @worksheet.tab_color = 'red'
    @worksheet.__send__('write_sheet_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>'
    assert_equal(expected, result)
  end

  def test_write_sheet_pr_fit_page_and_tab_color
    @worksheet.instance_variable_get(:@page_setup).fit_page = true
    @worksheet.tab_color = 'red'
    @worksheet.__send__('write_sheet_pr')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>'
    assert_equal(expected, result)
  end
end
