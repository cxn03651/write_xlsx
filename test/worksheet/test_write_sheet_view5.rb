# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView5 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views_set_selection_A1
    @worksheet.select
    @worksheet.set_selection('A1')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_A2
    @worksheet.select
    @worksheet.set_selection('A2')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_B1
    @worksheet.select
    @worksheet.set_selection('B1')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_D3
    @worksheet.select
    @worksheet.set_selection('D3')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_D3_F4
    @worksheet.select
    @worksheet.set_selection('D3:F4')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="D3" sqref="D3:F4"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_F4_D3
    @worksheet.select
    @worksheet.set_selection('F4:D3')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="F4" sqref="D3:F4"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_A2_A2
    @worksheet.select
    @worksheet.set_selection('A2:A2')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end
end
