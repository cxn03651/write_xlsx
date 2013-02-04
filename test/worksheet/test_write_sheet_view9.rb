# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView9 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views_set_selection_A2_split_panes_15_0_20_0
    @worksheet.select
    @worksheet.set_selection('A2')
    @worksheet.split_panes(15, 0, 20, 0)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A21" activePane="bottomLeft"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_A21_split_panes_15_0_20_0
    @worksheet.select
    @worksheet.set_selection('A21')
    @worksheet.split_panes(15, 0, 20, 0)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A21" activePane="bottomLeft"/><selection pane="bottomLeft" activeCell="A21" sqref="A21"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_B1_split_panes_0_843_0_4
    @worksheet.select
    @worksheet.set_selection('B1')
    @worksheet.split_panes(0, 8.43, 0, 4)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="E1" activePane="topRight"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_E1_split_panes_0_843_0_4
    @worksheet.select
    @worksheet.set_selection('E1')
    @worksheet.split_panes(0, 8.43, 0, 4)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="E1" activePane="topRight"/><selection pane="topRight" activeCell="E1" sqref="E1"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end
end
