# -*- coding: utf-8 -*-
require 'helper'
require 'write_xlsx'
require 'stringio'

class TestWriteSheetProtection < Test::Unit::TestCase
  def setup
    @workbook = WriteXLSX.new(StringIO.new)
    @worksheet = @workbook.add_worksheet('')
  end

  def test_write_sheet_protection_without_password
    @worksheet.protect
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_password
    @worksheet.protect('password')
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection password="83AF" sheet="1" objects="1" scenarios="1"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_select_locked_cells
    @worksheet.protect('', :select_locked_cells => false)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" selectLockedCells="1"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_format_cells
    @worksheet.protect('', :format_cells => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_format_columns
    @worksheet.protect('', :format_columns => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_format_rows
    @worksheet.protect('', :format_rows => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatRows="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_insert_columns
    @worksheet.protect('', :insert_columns => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertColumns="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_insert_rows
    @worksheet.protect('', :insert_rows => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertRows="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_insert_hyperlinks
    @worksheet.protect('', :insert_hyperlinks => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" insertHyperlinks="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_delete_columns
    @worksheet.protect('', :delete_columns => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" deleteColumns="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_delete_rows
    @worksheet.protect('', :delete_rows => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" deleteRows="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_sort
    @worksheet.protect('', :sort => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" sort="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_autofilter
    @worksheet.protect('', :autofilter => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" autoFilter="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_pivot_table
    @worksheet.protect('', :pivot_tables => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" pivotTables="0"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_objects
    @worksheet.protect('', :objects => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" scenarios="1"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_scenarios
    @worksheet.protect('', :scenarios => true)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_format_cells_and_select_locked_cells_and_select_unlocked_cells
    @worksheet.protect('', :format_cells => true,
             :select_locked_cells => false, :select_unlocked_cells => false)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1" selectUnlockedCells="1"/>'
    assert_equal(expected, result)
  end

  def test_write_sheet_protection_with_all
    password = 'drowssap'
    options = {
      :objects               => true,
      :scenarios             => true,
      :format_cells          => true,
      :format_columns        => true,
      :format_rows           => true,
      :insert_columns        => true,
      :insert_rows           => true,
      :insert_hyperlinks     => true,
      :delete_columns        => true,
      :delete_rows           => true,
      :select_locked_cells   => false,
      :sort                  => true,
      :autofilter            => true,
      :pivot_tables          => true,
      :select_unlocked_cells => false
    }
    @worksheet.protect(password, options)
    @worksheet.__send__('write_sheet_protection')
    result = @worksheet.instance_variable_get(:@writer).string
    expected = '<sheetProtection password="996B" sheet="1" formatCells="0" formatColumns="0" formatRows="0" insertColumns="0" insertRows="0" insertHyperlinks="0" deleteColumns="0" deleteRows="0" selectLockedCells="1" sort="0" autoFilter="0" pivotTables="0" selectUnlockedCells="1"/>'
    assert_equal(expected, result)
  end
end
