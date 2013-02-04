# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView8 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views_set_selection_A2_split_panes_15_0
    @worksheet.select
    @worksheet.set_selection('A2')
    @worksheet.split_panes(15, 0)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="600" topLeftCell="A2" activePane="bottomLeft"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_B1_split_panes_0_843
    @worksheet.select
    @worksheet.set_selection('B1')
    @worksheet.split_panes(0, 8.43)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="1350" topLeftCell="B1" activePane="topRight"/><selection pane="topRight" activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_G4_split_panes_45_5414
    @worksheet.select
    @worksheet.set_selection('G4')
    @worksheet.split_panes(45, 54.14)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" activePane="bottomRight"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="G4" sqref="G4"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_selection_I5_split_panes_45_5414
    @worksheet.select
    @worksheet.set_selection('I5')
    @worksheet.split_panes(45, 54.14)
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane xSplit="6150" ySplit="1200" topLeftCell="G4" activePane="bottomRight"/><selection pane="topRight" activeCell="G1" sqref="G1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="I5" sqref="I5"/></sheetView></sheetViews>'
    assert_equal(expected, result)
  end
end
