# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView2 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views_freeze_panes_1
    @worksheet.select
    @worksheet.freeze_panes(1)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_freeze_panes_0_1
    @worksheet.select
    @worksheet.freeze_panes(0, 1)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_freeze_panes_1_1
    @worksheet.select
    @worksheet.freeze_panes(1, 1)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_freeze_panes_G4
    @worksheet.select
    @worksheet.freeze_panes('G4')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_freeze_panes_3_6_3_6_1
    @worksheet.select
    @worksheet.freeze_panes(3, 6, 3, 6, 1)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozenSplit"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight"/></sheetView></sheetViews>';
    assert_equal(expected, result)
  end
end
