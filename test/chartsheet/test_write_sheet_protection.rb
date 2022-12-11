# -*- coding: utf-8 -*-

require 'helper'

class TestChartsheetWriteSheetProtection < Minitest::Test
  def setup
    workbook = WriteXLSX.new(StringIO.new)
    workbook.add_chart(:type => 'line')
    @chartsheet = workbook.sheets.first
  end

  def test_chartsheet_write_sheet_protection_with_blank_password
    expected = '<sheetProtection content="1" objects="1"/>'

    @chartsheet.protect('', {})
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_with_password
    expected = '<sheetProtection password="83AF" content="1" objects="1"/>'

    @chartsheet.protect('password', {})
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_without_password_and_content
    expected = '<sheetProtection content="1"/>'

    @chartsheet.protect('', { :objects => 0 })
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_with_opjects0_option
    expected = '<sheetProtection password="83AF" content="1"/>'

    @chartsheet.protect('password', { :objects => 0 })
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_without_password
    expected = '<sheetProtection objects="1"/>'

    @chartsheet.protect('', { :content => 0 })
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_without_password_and_content_option
    expected = ''

    @chartsheet.protect('', { :content => 0, :objects => 0 })
    result = @chartsheet.__send__(:write_sheet_protection) || ''

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_with_password_and_content_objects_option
    expected = '<sheetProtection password="83AF"/>'

    @chartsheet.protect('password', { :content => 0, :objects => 0 })
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end

  def test_chartsheet_write_sheet_protection_with_password_full_options
    expected = '<sheetProtection password="83AF" content="1" objects="1"/>'

    options = {
      :objects               => 1,
      :scenarios             => 1,
      :format_cells          => 1,
      :format_columns        => 1,
      :format_rows           => 1,
      :insert_columns        => 1,
      :insert_rows           => 1,
      :insert_hyperlinks     => 1,
      :delete_columns        => 1,
      :delete_rows           => 1,
      :select_locked_cells   => 0,
      :sort                  => 1,
      :autofilter            => 1,
      :pivot_tables          => 1,
      :select_unlocked_cells => 0
    }
    @chartsheet.protect('password', options)
    result = @chartsheet.__send__(:write_sheet_protection)

    assert_equal(expected_to_array(expected), got_to_array(result))
  end
end
