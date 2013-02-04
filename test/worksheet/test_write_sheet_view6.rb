# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView6 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views_set_selection_A2_freeze_panes_1_0
    @worksheet.select
    @worksheet.set_selection('A2')
    @worksheet.freeze_panes(1, 0)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_B1_freeze_panes_0_1
    @worksheet.select
    @worksheet.set_selection('B1')
    @worksheet.freeze_panes(0, 1)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_G4_freeze_panes_G4
    @worksheet.select
    @worksheet.set_selection('G4')
    @worksheet.freeze_panes('G4')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_I5_freeze_panes_G4
    @worksheet.select
    @worksheet.set_selection('I5')
    @worksheet.freeze_panes('G4')
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6" ySplit="3" topLeftCell="G4" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="I5" sqref="I5"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end
end
