# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView1 < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_views
    @worksheet.select
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_and_zoom_100
    @worksheet.select
    @worksheet.zoom = 100
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_and_zoom_200
    @worksheet.select
    @worksheet.zoom = 200
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" zoomScale="200" zoomScaleNormal="200" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_right_to_left
    @worksheet.select
    @worksheet.right_to_left
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_hide_zero
    @worksheet.select
    @worksheet.hide_zero
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView showZeros="0" tabSelected="1" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end

  def test_write_sheet_views_set_page_view
    @worksheet.select
    @worksheet.set_page_view
    @worksheet.__send__('write_sheet_views')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetViews><sheetView tabSelected="1" view="pageLayout" workbookViewId="0"/></sheetViews>'
    assert_equal(expected, result)
  end
end
