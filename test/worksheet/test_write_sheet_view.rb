# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetView < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_view_tab_not_selected
    @worksheet.__send__('write_sheet_view')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetView workbookViewId="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_view_tab_selected
    @worksheet.select
    @worksheet.__send__('write_sheet_view')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetView tabSelected="1" workbookViewId="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_view_tab_selected_and_hide_gridlines
    @worksheet.select
    @worksheet.hide_gridlines
    @worksheet.__send__('write_sheet_view')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetView tabSelected="1" workbookViewId="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_view_tab_selected_and_hide_gridlines_false
    @worksheet.select
    @worksheet.hide_gridlines(false)
    @worksheet.__send__('write_sheet_view')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetView tabSelected="1" workbookViewId="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_view_tab_selected_and_hide_gridlines_another
    @worksheet.select
    @worksheet.hide_gridlines(2)
    @worksheet.__send__('write_sheet_view')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetView showGridLines="0" tabSelected="1" workbookViewId="0"/>'
    assert_equal(expected, result)
  end
end
