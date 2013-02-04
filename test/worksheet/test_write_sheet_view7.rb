# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView7 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views_set_selection_A2_freeze_panes_1_0_20_0
    @worksheet.select
    @worksheet.set_selection('A2')
    @worksheet.freeze_panes(1, 0, 20, 0)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A21" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_A1_freeze_panes_1_0_20_0
    @worksheet.select
    @worksheet.set_selection('A1')
    @worksheet.freeze_panes(1, 0, 20, 0)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A21" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_B1_freeze_panes_0_1_0_4
    @worksheet.select
    @worksheet.set_selection('B1')
    @worksheet.freeze_panes(0, 1, 0, 4)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="E1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_A1_freeze_panes_0_1_0_4
    @worksheet.select
    @worksheet.set_selection('A1')
    @worksheet.freeze_panes(0, 1, 0, 4)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="E1" activePane="topRight" state="frozen"/><selection pane="topRight"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_G4_freeze_panes_3_6_6_8
    @worksheet.select
    @worksheet.set_selection('G4')
    @worksheet.freeze_panes(3, 6, 6, 8)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="I7" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_A1_freeze_panes_3_6_6_8
    @worksheet.select
    @worksheet.set_selection('A1')
    @worksheet.freeze_panes(3, 6, 6, 8)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="I7" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end
end
